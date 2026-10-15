import { useCallback, useEffect, useRef, useState, type PointerEvent } from "react";
import JSZip from "jszip";
import { createPortal } from "react-dom";
import { renderAsync } from "docx-preview";

const signers = [
  {
    role: "渠道方 · 盖章/签字",
    name: "河北汉恒紧固件有限公司",
    file: "渠道方_河北汉恒紧固件有限公司.png",
  },
  {
    role: "渠道变更方 · 盖章/签字",
    name: "河北曜境智能科技有限公司",
    file: "渠道变更方_河北曜境智能科技有限公司.png",
  },
];

const SIGNING_STORAGE_KEY = "contract-signing-v1";

type SavedSigningData = {
  signatures: (string | null)[];
  positions: { x: number; y: number }[];
  history: SignatureHistoryItem[];
};

type SignatureHistoryItem = {
  id: string;
  signerIndex: number;
  signature: string;
  createdAt: string;
};

function normalizeSigningData(value: Partial<SavedSigningData>): SavedSigningData | null {
  if (!Array.isArray(value.signatures) || !Array.isArray(value.positions)) return null;
  const signatures = signers.map((_signer, index) =>
    typeof value.signatures?.[index] === "string" ? value.signatures[index] as string : null);
  const positions = signers.map((_signer, index) => {
    const position = value.positions?.[index];
    if (!position || typeof position.x !== "number" || typeof position.y !== "number") return { x: 0.4, y: 0.43 };
    return { x: Math.min(1, Math.max(0, position.x)), y: Math.min(1, Math.max(0, position.y)) };
  });
  const history = Array.isArray(value.history)
    ? value.history.filter((item): item is SignatureHistoryItem =>
      Boolean(item) && typeof item.id === "string" && Number.isInteger(item.signerIndex) &&
      item.signerIndex >= 0 && item.signerIndex < signers.length &&
      typeof item.signature === "string" && item.signature.startsWith("data:image/png;base64,") &&
      typeof item.createdAt === "string")
    : signatures.flatMap((signature, signerIndex) => signature ? [{
      id: `legacy-${signerIndex}`, signerIndex, signature, createdAt: new Date().toISOString(),
    }] : []);
  return { signatures, positions, history };
}

function readSavedSigningData(): SavedSigningData | null {
  try {
    const raw = localStorage.getItem(SIGNING_STORAGE_KEY);
    if (!raw) return null;
    return normalizeSigningData(JSON.parse(raw) as Partial<SavedSigningData>);
  } catch {
    return null;
  }
}

export default function SignPage() {
  const [initialSavedData] = useState(readSavedSigningData);
  const hasStoredPositionsRef = useRef(Boolean(initialSavedData));
  const contractRef = useRef<HTMLDivElement>(null);
  const [signatureTargets, setSignatureTargets] = useState<{ page: HTMLElement; x: number; y: number }[]>([]);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const drawingRef = useRef(false);
  const [signerIndex, setSignerIndex] = useState(0);
  const [hasInk, setHasInk] = useState(false);
  const [message, setMessage] = useState("");
  const [isDownloading, setIsDownloading] = useState(false);
  const [signatures, setSignatures] = useState<(string | null)[]>(initialSavedData?.signatures ?? [null, null]);
  const [signaturePositions, setSignaturePositions] = useState(initialSavedData?.positions ?? [{ x: 0.40, y: 0.43 }, { x: 0.40, y: 0.43 }]);
  const [signatureHistory, setSignatureHistory] = useState<SignatureHistoryItem[]>(initialSavedData?.history ?? []);
  const [contractReady, setContractReady] = useState(false);
  const [serverReady, setServerReady] = useState(false);
  const [canvasRevision, setCanvasRevision] = useState(0);
  const dragRef = useRef<null | {
    signerIndex: number;
    startX: number;
    startY: number;
    startPositionX: number;
    startPositionY: number;
    width: number;
    height: number;
    imageWidth: number;
    imageHeight: number;
  }>(null);

  useEffect(() => {
    let cancelled = false;
    fetch("/api/signing", { headers: { Accept: "application/json" } })
      .then(async (response) => {
        if (!response.ok) throw new Error("服务器签名数据读取失败");
        return response.json() as Promise<{ data: Partial<SavedSigningData> | null }>;
      })
      .then(({ data }) => {
        if (cancelled) return;
        const normalized = data ? normalizeSigningData(data) : null;
        if (normalized) {
          setSignatures(normalized.signatures);
          setSignaturePositions(normalized.positions);
          setSignatureHistory(normalized.history);
          hasStoredPositionsRef.current = true;
        }
        setServerReady(true);
      })
      .catch(() => {
        if (!cancelled) {
          setMessage("暂时无法连接服务器，签名不会同步到其他设备。");
          setServerReady(true);
        }
      });
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!contractReady || !serverReady) return;
    const timer = window.setTimeout(async () => {
      try {
        try {
          localStorage.setItem(SIGNING_STORAGE_KEY, JSON.stringify({ signatures, positions: signaturePositions, history: signatureHistory }));
        } catch {
          // Server persistence remains the primary save path if browser storage is unavailable.
        }
        const response = await fetch("/api/signing", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ signatures, positions: signaturePositions, history: signatureHistory }),
        });
        if (!response.ok) throw new Error("服务器未能保存签名");
        setMessage(signatures.some(Boolean)
          ? "签名和位置已保存到服务器，关闭浏览器后仍可继续查看。"
          : "服务器连接正常，完成签字后会自动保存。" );
      } catch {
        setMessage("服务器保存失败；签名仅暂存在本机，请检查网络后重试。");
      }
    }, 500);
    return () => window.clearTimeout(timer);
  }, [contractReady, serverReady, signatures, signaturePositions, signatureHistory]);

  /*
   * The contract preview and signature editor are local, but persisted signing data
   * is shared through the Netlify Function after the initial server read completes.
   */
  useEffect(() => {
    let cancelled = false;
    const host = contractRef.current;
    if (!host) return;
    fetch("/contract-template.docx")
      .then((response) => {
        if (!response.ok) throw new Error("找不到原始 Word 合同文件");
        return response.arrayBuffer();
      })
      .then(async (buffer) => {
        host.replaceChildren();
        await renderAsync(buffer, host, host, { breakPages: true, ignoreLastRenderedPageBreak: false });
        if (cancelled) return;
        const pages = Array.from(host.querySelectorAll<HTMLElement>(".docx"));
        const paragraphs = Array.from(host.querySelectorAll<HTMLElement>("p"));
        const signatureParagraphs = paragraphs.filter((p) => (p.innerText || "").includes("签字") && /_{6,}/.test(p.innerText || ""));
        const targets = signers.map((_signer, index) => {
          const paragraph = signatureParagraphs[index];
          if (!paragraph) return null;
          const page = paragraph.closest<HTMLElement>(".docx") || pages[0];
          if (!page) return null;
          page.style.position = "relative";
          const pageRect = page.getBoundingClientRect();
          const pageWidth = page.clientWidth || pageRect.width;
          const pageHeight = page.clientHeight || pageRect.height;
          const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
          let textNode = walker.nextNode();
          while (textNode && !/_{6,}/.test(textNode.textContent || "")) textNode = walker.nextNode();
          let x = 0.4;
          let y = 0.43;
          if (textNode) {
            const range = document.createRange();
            const lineText = textNode.textContent || "";
            const lineStart = lineText.search(/_{6,}/);
            const lineEnd = lineText.indexOf(" ", lineStart);
            range.setStart(textNode, lineStart);
            range.setEnd(textNode, lineEnd > lineStart ? lineEnd : lineText.length);
            const rect = range.getBoundingClientRect();
            if (pageWidth && rect.width) x = (rect.left - pageRect.left - page.clientLeft) / pageWidth;
            if (pageHeight && rect.height) y = (rect.top - pageRect.top - page.clientTop - 21) / pageHeight;
          }
          return { page, x: Math.min(.88, Math.max(0, x)), y: Math.min(.92, Math.max(0, y)) };
        });
        const resolved = targets.filter((item): item is { page: HTMLElement; x: number; y: number } => item !== null);
        setSignatureTargets(resolved);
        if (!hasStoredPositionsRef.current) {
          setSignaturePositions(resolved.map(({ x, y }) => ({ x, y })));
        }
        setContractReady(true);
      })
      .catch((error) => { if (!cancelled) setMessage(error instanceof Error ? error.message : "合同预览加载失败"); });
    return () => { cancelled = true; };
  }, []);

  const prepareCanvas = useCallback(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const bounds = canvas.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    let previousCanvas: HTMLCanvasElement | null = null;
    if (canvas.width > 0 && canvas.height > 0) {
      previousCanvas = document.createElement("canvas");
      previousCanvas.width = canvas.width;
      previousCanvas.height = canvas.height;
      previousCanvas.getContext("2d")?.drawImage(canvas, 0, 0);
    }
    canvas.width = Math.round(bounds.width * dpr);
    canvas.height = Math.round(bounds.height * dpr);
    const context = canvas.getContext("2d");
    if (!context) return;
    context.scale(dpr, dpr);
    context.lineWidth = 3.2;
    context.lineCap = "round";
    context.lineJoin = "round";
    context.strokeStyle = "#203d35";
    if (previousCanvas) context.drawImage(previousCanvas, 0, 0, bounds.width, bounds.height);
  }, []);

  useEffect(() => {
    prepareCanvas();
    const canvas = canvasRef.current;
    if (!canvas) return;
    const observer = new ResizeObserver(() => {
      prepareCanvas();
    });
    observer.observe(canvas);
    return () => observer.disconnect();
  }, [prepareCanvas]);

  useEffect(() => {
    if (!serverReady) return;
    const canvas = canvasRef.current;
    const savedSignature = signatures[signerIndex];
    const context = canvas?.getContext("2d");
    if (!canvas || !context) return;
    if (!savedSignature) {
      context.clearRect(0, 0, canvas.width, canvas.height);
      setHasInk(false);
      return;
    }
    const image = new Image();
    image.onload = () => {
      context.clearRect(0, 0, canvas.width, canvas.height);
      context.drawImage(image, 0, 0, canvas.clientWidth, canvas.clientHeight);
      setHasInk(true);
    };
    image.src = savedSignature;
  }, [signerIndex, serverReady, canvasRevision]);

  const pointFromEvent = (event: PointerEvent<HTMLCanvasElement>) => {
    const rect = event.currentTarget.getBoundingClientRect();
    return { x: event.clientX - rect.left, y: event.clientY - rect.top };
  };

  const startDrawing = (event: PointerEvent<HTMLCanvasElement>) => {
    if (event.pointerType === "mouse" && event.button !== 0) return;
    const canvas = event.currentTarget;
    const context = canvas.getContext("2d");
    if (!context) return;
    event.preventDefault();
    canvas.setPointerCapture(event.pointerId);
    const point = pointFromEvent(event);
    drawingRef.current = true;
    context.beginPath();
    context.moveTo(point.x, point.y);
    context.lineTo(point.x + 0.1, point.y + 0.1);
    context.stroke();
    setHasInk(true);
    setSignatures((current) => current[signerIndex]
      ? current.map((signature, index) => index === signerIndex ? null : signature)
      : current);
    setMessage("");
  };

  const continueDrawing = (event: PointerEvent<HTMLCanvasElement>) => {
    if (!drawingRef.current) return;
    const context = event.currentTarget.getContext("2d");
    if (!context) return;
    event.preventDefault();
    const point = pointFromEvent(event);
    context.lineTo(point.x, point.y);
    context.stroke();
  };

  const stopDrawing = () => {
    drawingRef.current = false;
  };

  const clearSignature = () => {
    const canvas = canvasRef.current;
    const context = canvas?.getContext("2d");
    if (canvas && context) context.clearRect(0, 0, canvas.width, canvas.height);
    setHasInk(false);
    setSignatures((current) => current.map((signature, index) => index === signerIndex ? null : signature));
    setMessage("");
  };

  const selectSigner = (index: number) => {
    if (index === signerIndex) return;
    if (hasInk && !signatures[signerIndex] && !window.confirm("切换签字方会清除尚未完成的签名，确定继续吗？")) return;
    const canvas = canvasRef.current;
    const context = canvas?.getContext("2d");
    if (canvas && context) context.clearRect(0, 0, canvas.width, canvas.height);
    setHasInk(false);
    setMessage("");
    setSignerIndex(index);
  };

  const completeSignature = () => {
    const canvas = canvasRef.current;
    if (!canvas || !hasInk) return;
    const signature = canvas.toDataURL("image/png");
    const historyItem: SignatureHistoryItem = {
      id: typeof crypto !== "undefined" && "randomUUID" in crypto
        ? crypto.randomUUID()
        : `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      signerIndex,
      signature,
      createdAt: new Date().toISOString(),
    };
    setSignatures((current) => current.map((signature, index) =>
      index === signerIndex ? historyItem.signature : signature));
    setSignatureHistory((current) => [...current, historyItem]);
    setMessage("签名已放到左侧合同签字栏，可拖动签名调整位置。");
  };

  const deleteHistoryItem = (item: SignatureHistoryItem) => {
    const nextHistory = signatureHistory.filter((saved) => saved.id !== item.id);
    setSignatureHistory(nextHistory);
    setMessage("历史签名已删除。");
    if (signatures[item.signerIndex] !== item.signature) return;

    const replacement = [...nextHistory].reverse().find((saved) => saved.signerIndex === item.signerIndex);
    setSignatures((current) => current.map((signature, index) =>
      index === item.signerIndex ? replacement?.signature ?? null : signature));
    if (signerIndex === item.signerIndex) {
      const canvas = canvasRef.current;
      const context = canvas?.getContext("2d");
      if (!canvas || !context) return;
      context.clearRect(0, 0, canvas.width, canvas.height);
      if (replacement) {
        const image = new Image();
        image.onload = () => {
          context.clearRect(0, 0, canvas.width, canvas.height);
          context.drawImage(image, 0, 0, canvas.clientWidth, canvas.clientHeight);
          setHasInk(true);
        };
        image.src = replacement.signature;
      } else {
        setHasInk(false);
      }
    }
    setMessage("历史签名已删除，并已同步更新合同签名。");
  };

  const reuseHistoryItem = (item: SignatureHistoryItem) => {
    if (hasInk && !signatures[signerIndex] &&
        !window.confirm("复用历史签名会清除尚未完成的手写内容，确定继续吗？")) return;
    setSignerIndex(item.signerIndex);
    setSignatures((current) => current.map((signature, index) =>
      index === item.signerIndex ? item.signature : signature));
    setCanvasRevision((current) => current + 1);
    setMessage(`${signers[item.signerIndex].name}的历史签名已复用到合同。`);
  };

  const startMovingSignature = (event: PointerEvent<HTMLImageElement>, index: number) => {
    const area = event.currentTarget.parentElement;
    if (!area) return;
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.setPointerCapture(event.pointerId);
    dragRef.current = {
      signerIndex: index,
      startX: event.clientX,
      startY: event.clientY,
      startPositionX: signaturePositions[index].x,
      startPositionY: signaturePositions[index].y,
      width: area.getBoundingClientRect().width,
      height: area.getBoundingClientRect().height,
      imageWidth: event.currentTarget.getBoundingClientRect().width,
      imageHeight: event.currentTarget.getBoundingClientRect().height,
    };
  };

  const moveSignature = (event: PointerEvent<HTMLImageElement>) => {
    const drag = dragRef.current;
    if (!drag) return;
    event.preventDefault();
    const maxX = Math.max(0, 1 - drag.imageWidth / drag.width);
    const maxY = Math.max(0, 1 - drag.imageHeight / drag.height);
    const nextX = Math.min(maxX, Math.max(0,
      drag.startPositionX + (event.clientX - drag.startX) / drag.width));
    const nextY = Math.min(maxY, Math.max(0,
      drag.startPositionY + (event.clientY - drag.startY) / drag.height));
    setSignaturePositions((current) => current.map((position, index) =>
      index === drag.signerIndex ? { x: nextX, y: nextY } : position));
  };

  const stopMovingSignature = () => { dragRef.current = null; };

  const downloadContract = async () => {
    const signedItems = signatures
      .map((dataUrl, index) => dataUrl ? { dataUrl, signerIndex: index } : null)
      .filter((item): item is { dataUrl: string; signerIndex: number } => item !== null);
    if (!signedItems.length) return;
    setIsDownloading(true);
    setMessage("");
    try {
      const response = await fetch("/contract-template.docx");
      if (!response.ok) throw new Error("找不到原始 Word 合同文件");
      const zip = await JSZip.loadAsync(await response.arrayBuffer());
      const documentFile = zip.file("word/document.xml");
      const relsFile = zip.file("word/_rels/document.xml.rels");
      const typesFile = zip.file("[Content_Types].xml");
      if (!documentFile || !relsFile || !typesFile) throw new Error("原始合同文件结构不完整");

      const [documentText, relsText, typesText] = await Promise.all([
        documentFile.async("text"),
        relsFile.async("text"),
        typesFile.async("text"),
      ]);
      const parser = new DOMParser();
      const documentXml = parser.parseFromString(documentText, "application/xml");
      const relsXml = parser.parseFromString(relsText, "application/xml");
      const typesXml = parser.parseFromString(typesText, "application/xml");
      if ([documentXml, relsXml, typesXml].some((xml) => xml.querySelector("parsererror"))) {
        throw new Error("无法解析原始合同文件");
      }

      const wordNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
      const relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
      const pkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
      const wpNs = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
      const drawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
      const pictureNs = "http://schemas.openxmlformats.org/drawingml/2006/picture";
      const paragraphs = Array.from(documentXml.getElementsByTagNameNS(wordNs, "p"));
      const relationNodes = Array.from(relsXml.getElementsByTagNameNS(pkgRelNs, "Relationship"));
      const usedIds = relationNodes.map((node) => node.getAttribute("Id") || "");
      let relIdNumber = 1;
      while (usedIds.includes(`rId${relIdNumber}`)) relIdNumber += 1;
      const contentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
      const extentWidth = 822_960; // 0.9 in, in EMU
      const aspect = 2.35;
      const extentHeight = Math.round(extentWidth / aspect);
      const existingPngType = Array.from(typesXml.getElementsByTagNameNS(contentTypesNs, "Default"))
        .some((node) => node.getAttribute("Extension")?.toLowerCase() === "png");
      if (!existingPngType) {
        const pngType = typesXml.createElementNS(contentTypesNs, "Default");
        pngType.setAttribute("Extension", "png"); pngType.setAttribute("ContentType", "image/png");
        typesXml.documentElement.appendChild(pngType);
      }

      for (const item of signedItems) {
      const selected = signers[item.signerIndex];
      const targetParagraph = paragraphs.find((paragraph) => {
        const text = paragraph.textContent || "";
        return text.includes(selected.name) && text.includes("法定代表人/负责人/本人签字");
      });
      if (!targetParagraph) throw new Error(`找不到${selected.name}的签字栏`);
      const signatureText = Array.from(targetParagraph.getElementsByTagNameNS(wordNs, "t"))
        .find((node) => /_{6,}/.test(node.textContent || ""));
      if (!signatureText) throw new Error(`找不到${selected.name}的签名位置`);
      const placeholderRun = signatureText.parentElement;
      if (!placeholderRun) throw new Error("无法定位合同签字位置");

      const relationId = `rId${relIdNumber++}`;
      const relation = relsXml.createElementNS(pkgRelNs, "Relationship");
      relation.setAttribute("Id", relationId);
      relation.setAttribute("Type", `${relNs}/image`);
      relation.setAttribute("Target", `media/signature-${item.signerIndex}.png`);
      relsXml.documentElement.appendChild(relation);
      const imageBytes = await (await fetch(item.dataUrl)).arrayBuffer();
      const nextDrawingId = Math.max(0, ...Array.from(documentXml.getElementsByTagNameNS(wpNs, "docPr"))
        .map((node) => Number(node.getAttribute("id")) || 0)) + 1;
      const drawingRun = documentXml.createElementNS(wordNs, "w:r");
      const drawing = documentXml.createElementNS(wordNs, "w:drawing");
      const anchor = documentXml.createElementNS(wpNs, "wp:anchor");
      anchor.setAttribute("distT", "0"); anchor.setAttribute("distB", "0");
      anchor.setAttribute("distL", "0"); anchor.setAttribute("distR", "0");
      anchor.setAttribute("simplePos", "0"); anchor.setAttribute("relativeHeight", "251659264");
      anchor.setAttribute("behindDoc", "0"); anchor.setAttribute("locked", "0");
      anchor.setAttribute("layoutInCell", "1"); anchor.setAttribute("allowOverlap", "1");
      const simplePosition = documentXml.createElementNS(wpNs, "wp:simplePos");
      simplePosition.setAttribute("x", "0"); simplePosition.setAttribute("y", "0");
      anchor.appendChild(simplePosition);
      const horizontalPosition = documentXml.createElementNS(wpNs, "wp:positionH");
      horizontalPosition.setAttribute("relativeFrom", "page");
      const horizontalOffset = documentXml.createElementNS(wpNs, "wp:posOffset");
      horizontalOffset.textContent = String(Math.round(signaturePositions[item.signerIndex].x * 7_561_200));
      horizontalPosition.appendChild(horizontalOffset);
      const verticalPositionAnchor = documentXml.createElementNS(wpNs, "wp:positionV");
      verticalPositionAnchor.setAttribute("relativeFrom", "page");
      const verticalOffsetAnchor = documentXml.createElementNS(wpNs, "wp:posOffset");
      verticalOffsetAnchor.textContent = String(Math.round(signaturePositions[item.signerIndex].y * 10_693_650));
      verticalPositionAnchor.appendChild(verticalOffsetAnchor);
      anchor.appendChild(horizontalPosition); anchor.appendChild(verticalPositionAnchor);
      const extent = documentXml.createElementNS(wpNs, "wp:extent");
      extent.setAttribute("cx", String(extentWidth)); extent.setAttribute("cy", String(extentHeight));
      anchor.appendChild(extent);
      const effectExtent = documentXml.createElementNS(wpNs, "wp:effectExtent");
      effectExtent.setAttribute("l", "0"); effectExtent.setAttribute("t", "0");
      effectExtent.setAttribute("r", "0"); effectExtent.setAttribute("b", "0");
      anchor.appendChild(effectExtent);
      anchor.appendChild(documentXml.createElementNS(wpNs, "wp:wrapNone"));
      const docPr = documentXml.createElementNS(wpNs, "wp:docPr");
      docPr.setAttribute("id", String(nextDrawingId)); docPr.setAttribute("name", `签名 - ${selected.name}`);
      anchor.appendChild(docPr);
      const frame = documentXml.createElementNS(wpNs, "wp:cNvGraphicFramePr");
      const frameLocks = documentXml.createElementNS(drawingNs, "a:graphicFrameLocks");
      frameLocks.setAttribute("noChangeAspect", "1"); frame.appendChild(frameLocks);
      anchor.appendChild(frame);
      const graphic = documentXml.createElementNS(drawingNs, "a:graphic");
      const graphicData = documentXml.createElementNS(drawingNs, "a:graphicData");
      graphicData.setAttribute("uri", pictureNs);
      const picture = documentXml.createElementNS(pictureNs, "pic:pic");
      const nonVisual = documentXml.createElementNS(pictureNs, "pic:nvPicPr");
      const pictureProperties = documentXml.createElementNS(pictureNs, "pic:cNvPr");
      pictureProperties.setAttribute("id", "0"); pictureProperties.setAttribute("name", `signature-${item.signerIndex}.png`);
      nonVisual.appendChild(pictureProperties);
      nonVisual.appendChild(documentXml.createElementNS(pictureNs, "pic:cNvPicPr"));
      const fill = documentXml.createElementNS(pictureNs, "pic:blipFill");
      const blip = documentXml.createElementNS(drawingNs, "a:blip");
      blip.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:r", relNs);
      blip.setAttributeNS(relNs, "r:embed", relationId);
      fill.appendChild(blip);
      const stretch = documentXml.createElementNS(drawingNs, "a:stretch");
      stretch.appendChild(documentXml.createElementNS(drawingNs, "a:fillRect"));
      fill.appendChild(stretch);
      const shapeProperties = documentXml.createElementNS(pictureNs, "pic:spPr");
      const transform = documentXml.createElementNS(drawingNs, "a:xfrm");
      const offset = documentXml.createElementNS(drawingNs, "a:off");
      offset.setAttribute("x", "0"); offset.setAttribute("y", "0");
      const shapeExtent = documentXml.createElementNS(drawingNs, "a:ext");
      shapeExtent.setAttribute("cx", String(extentWidth)); shapeExtent.setAttribute("cy", String(extentHeight));
      transform.appendChild(offset); transform.appendChild(shapeExtent);
      shapeProperties.appendChild(transform);
      const geometry = documentXml.createElementNS(drawingNs, "a:prstGeom");
      geometry.setAttribute("prst", "rect");
      geometry.appendChild(documentXml.createElementNS(drawingNs, "a:avLst"));
      shapeProperties.appendChild(geometry);
      picture.appendChild(nonVisual); picture.appendChild(fill); picture.appendChild(shapeProperties);
      graphicData.appendChild(picture); graphic.appendChild(graphicData);
      anchor.appendChild(graphic);
      drawing.appendChild(anchor); drawingRun.appendChild(drawing);
      placeholderRun.parentNode?.insertBefore(drawingRun, placeholderRun.nextSibling);
      zip.file(`word/media/signature-${item.signerIndex}.png`, imageBytes);
      }

      const serializer = new XMLSerializer();
      zip.file("word/document.xml", serializer.serializeToString(documentXml));
      zip.file("word/_rels/document.xml.rels", serializer.serializeToString(relsXml));
      zip.file("[Content_Types].xml", serializer.serializeToString(typesXml));
      const signedDocx = await zip.generateAsync({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      const url = URL.createObjectURL(signedDocx);
      const link = document.createElement("a");
      link.href = url;
      link.download = "签署版_渠道营销服务费收款账户变更证明.docx";
      link.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
      setMessage(`已下载包含 ${signedItems.length} 方签名的 Word 合同副本。`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "合同生成失败，请重试。");
    } finally {
      setIsDownloading(false);
    }
  };

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand-mark" aria-hidden="true">签</div>
        <div className="heading">
          <span className="eyebrow">CONTRACT SIGNING</span>
          <h1>账户变更证明签字</h1>
        </div>
        <span className="secure-label"><span />签名同步保存</span>
      </header>

      <div className="workspace">
        <section className="contract-card" aria-label="原始合同预览">
          <div className="document-label"><span className="document-icon">文</span>原始合同预览 <span className="label-divider">/</span> 内容与格式保持原样</div>
          <div className="contract-preview-scroll">
            <div className="contract-preview" ref={contractRef} />
            {signatures.map((signature, index) => signature && signatureTargets[index] && createPortal(
              <img
                key={`${signers[index].name}-signature`}
                className="signature-stamp"
                src={signature}
                alt={`${signers[index].name}的手写签名，可拖动调整位置`}
                draggable={false}
                style={{ left: `${signaturePositions[index].x * 100}%`, top: `${signaturePositions[index].y * 100}%` }}
                onPointerDown={(event) => startMovingSignature(event, index)}
                onPointerMove={moveSignature}
                onPointerUp={stopMovingSignature}
                onPointerCancel={stopMovingSignature}
                onLostPointerCapture={stopMovingSignature}
              />,
              signatureTargets[index].page,
            ))}
          </div>
        </section>

        <section className="signing-card" aria-labelledby="sign-heading">
          <div className="sign-card-heading">
            <div><span className="step-label">STEP 01</span><h2 id="sign-heading">选择签字方</h2></div>
            <span className="signing-hint">请核对你代表的公司</span>
          </div>
          <div className="party-options">
            {signers.map((signer, index) => (
              <button
                className={`party-option ${signerIndex === index ? "party-selected" : ""}`}
                key={signer.name}
                type="button"
                onClick={() => selectSigner(index)}
                aria-pressed={signerIndex === index}
              >
                <span className="radio">{signerIndex === index ? "✓" : ""}</span>
                <span className="party-copy"><span className="party-role">{signer.role}</span><strong>{signer.name}</strong></span>
              </button>
            ))}
          </div>

          <div className="signature-heading">
            <div><span className="step-label">STEP 02</span><h3>法定代表人 / 负责人 / 本人签字</h3></div>
            <span className="signature-caption">签字区域</span>
          </div>
          <div className={`signature-paper ${hasInk ? "has-ink" : ""}`}>
            <canvas
              ref={canvasRef}
              aria-label="手写签名区域"
              onPointerDown={startDrawing}
              onPointerMove={continueDrawing}
              onPointerUp={stopDrawing}
              onPointerCancel={stopDrawing}
              onLostPointerCapture={stopDrawing}
            />
            {!hasInk && <div className="placeholder"><span>请在此处手写签名</span><small>可使用鼠标或触屏书写</small></div>}
            <i className="corner corner-tl" /><i className="corner corner-tr" />
            <i className="corner corner-bl" /><i className="corner corner-br" />
          </div>
          <div className="privacy-note"><span className="tip-dot" />签字和位置自动同步到服务器，重新打开页面仍可查看</div>
          {message && <p className="status-message" role="status">{message}</p>}
          <div className="actions">
            <button className="button button-quiet" type="button" onClick={clearSignature} disabled={!hasInk && !signatures[signerIndex]}>重写当前签字</button>
            <button className="button button-primary" type="button" onClick={completeSignature} disabled={!hasInk || Boolean(signatures[signerIndex])}>
              签字完成，放入左侧 <span aria-hidden="true">←</span>
            </button>
          </div>
          <button className="button button-download" type="button" onClick={downloadContract} disabled={!signatures.some(Boolean) || isDownloading}>
            {isDownloading ? "正在生成合同…" : "下载合同（Word）"} <span aria-hidden="true">↓</span>
          </button>
          <section className="signature-vault" aria-labelledby="signature-vault-heading">
            <div className="vault-heading">
              <div>
                <span className="step-label">SIGNATURE RECORDS</span>
                <h3 id="signature-vault-heading">签名留存处</h3>
              </div>
              <span className="vault-count">{signatures.filter(Boolean).length} / {signers.length} 方已签</span>
            </div>
            <div className="vault-list">
              {signers.map((signer, index) => (
                <div className="vault-row" key={`vault-${signer.name}`}>
                  <div className="vault-party">
                    <span className={`vault-status ${signatures[index] ? "vault-status-done" : ""}`} aria-hidden="true">
                      {signatures[index] ? "✓" : "·"}
                    </span>
                    <div className="vault-copy">
                      <strong>{signer.name}</strong>
                      <span>{signatures[index] ? "已留存 · 自动同步" : "待签署"}</span>
                    </div>
                  </div>
                  {signatures[index] ? (
                    <img className="vault-signature" src={signatures[index] || undefined} alt={`${signer.name}已留存的签名`} />
                  ) : (
                    <span className="vault-empty">签名完成后显示</span>
                  )}
                </div>
              ))}
            </div>
            <div className="vault-history-heading">
              <strong>历史签名</strong>
              <span>{signatureHistory.length} 条</span>
            </div>
            {signatureHistory.length ? (
              <div className="vault-history-list">
                {[...signatureHistory].reverse().map((item) => (
                  <article className="vault-history-item" key={item.id}>
                    <div className="vault-history-info">
                      <strong>{signers[item.signerIndex].name}</strong>
                      <time dateTime={item.createdAt}>{new Date(item.createdAt).toLocaleString("zh-CN", { year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" })}</time>
                    </div>
                    <img className="vault-signature" src={item.signature} alt={`${signers[item.signerIndex].name}的历史签名`} />
                    <button
                      className="vault-reuse"
                      type="button"
                      onClick={() => reuseHistoryItem(item)}
                      aria-label={`复用 ${signers[item.signerIndex].name} 的这条历史签名`}
                    >
                      复用
                    </button>
                    <button
                      className="vault-delete"
                      type="button"
                      onClick={() => deleteHistoryItem(item)}
                      aria-label={`删除 ${signers[item.signerIndex].name} 的这条历史签名`}
                    >
                      删除
                    </button>
                  </article>
                ))}
              </div>
            ) : (
              <p className="vault-history-empty">点击“签字完成”后，签名记录会保存在这里。</p>
            )}
          </section>
          <p className="footer-note">会将已完成的签名放入合同对应签字栏；可分别完成双方签名后一次下载</p>
        </section>
      </div>

      <footer className="page-footer"><span>渠道营销服务费收款账户变更证明</span><span>签署前请确认账户信息与签字方</span></footer>
    </main>
  );
}

