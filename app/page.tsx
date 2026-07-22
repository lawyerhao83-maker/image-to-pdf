"use client";

import { useCallback, useRef, useState } from "react";
import {
  DndContext,
  DragEndEvent,
  PointerSensor,
  closestCenter,
  useSensor,
  useSensors,
} from "@dnd-kit/core";
import {
  SortableContext,
  arrayMove,
  rectSortingStrategy,
  useSortable,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import {
  ArrowDown,
  ArrowLeft,
  ArrowRight,
  ArrowUp,
  Check,
  Download,
  FileImage,
  GripVertical,
  ImagePlus,
  LoaderCircle,
  Lock,
  RotateCw,
  Trash2,
  UploadCloud,
  X,
} from "lucide-react";
import { jsPDF } from "jspdf";

type ImageItem = {
  id: string;
  file: File;
  url: string;
  width: number;
  height: number;
  rotation: number;
};

type SortableCardProps = {
  item: ImageItem;
  index: number;
  total: number;
  onRemove: (id: string) => void;
  onRotate: (id: string) => void;
  onMove: (from: number, to: number) => void;
};

function SortableCard({ item, index, total, onRemove, onRotate, onMove }: SortableCardProps) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({ id: item.id });

  return (
    <article
      ref={setNodeRef}
      style={{ transform: CSS.Transform.toString(transform), transition }}
      className={`image-card${isDragging ? " is-dragging" : ""}`}
    >
      <button className="drag-handle" type="button" aria-label={`拖曳第 ${index + 1} 張圖片`} {...attributes} {...listeners}>
        <GripVertical size={18} />
      </button>
      <div className="page-number">{String(index + 1).padStart(2, "0")}</div>
      <div className="thumb-wrap">
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img src={item.url} alt={item.file.name} style={{ transform: `rotate(${item.rotation}deg)` }} />
      </div>
      <div className="card-meta">
        <p title={item.file.name}>{item.file.name}</p>
        <span>{Math.round(item.file.size / 1024).toLocaleString()} KB</span>
      </div>
      <div className="card-actions">
        <button type="button" onClick={() => onMove(index, index - 1)} disabled={index === 0} aria-label="往前移">
          {index < 2 ? <ArrowLeft size={16} /> : <ArrowUp size={16} />}
        </button>
        <button type="button" onClick={() => onMove(index, index + 1)} disabled={index === total - 1} aria-label="往後移">
          {index < 2 ? <ArrowRight size={16} /> : <ArrowDown size={16} />}
        </button>
        <button type="button" onClick={() => onRotate(item.id)} aria-label="順時針旋轉">
          <RotateCw size={16} />
        </button>
        <button type="button" className="danger" onClick={() => onRemove(item.id)} aria-label="移除圖片">
          <Trash2 size={16} />
        </button>
      </div>
    </article>
  );
}

function loadImage(url: string) {
  return new Promise<HTMLImageElement>((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = reject;
    image.src = url;
  });
}

function renderRotatedImage(image: HTMLImageElement, rotation: number) {
  const sideways = rotation % 180 !== 0;
  const canvas = document.createElement("canvas");
  canvas.width = sideways ? image.naturalHeight : image.naturalWidth;
  canvas.height = sideways ? image.naturalWidth : image.naturalHeight;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("無法處理圖片");
  ctx.translate(canvas.width / 2, canvas.height / 2);
  ctx.rotate((rotation * Math.PI) / 180);
  ctx.drawImage(image, -image.naturalWidth / 2, -image.naturalHeight / 2);
  return canvas;
}

export default function Home() {
  const [images, setImages] = useState<ImageItem[]>([]);
  const [isOver, setIsOver] = useState(false);
  const [orientation, setOrientation] = useState<"portrait" | "landscape">("portrait");
  const [pageSize, setPageSize] = useState<"a4" | "letter" | "match">("a4");
  const [margin, setMargin] = useState<0 | 3 | 10>(3);
  const [fitMode, setFitMode] = useState<"contain" | "fill">("contain");
  const [fileName, setFileName] = useState("我的圖片合輯");
  const [isGenerating, setIsGenerating] = useState(false);
  const [notice, setNotice] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);
  const sensors = useSensors(useSensor(PointerSensor, { activationConstraint: { distance: 6 } }));

  const addFiles = useCallback(async (files: FileList | File[]) => {
    const accepted = Array.from(files).filter((file) => /image\/(jpeg|jpg)/i.test(file.type) || /\.jpe?g$/i.test(file.name));
    if (!accepted.length) {
      setNotice("請選擇 JPG 或 JPEG 圖片。");
      return;
    }
    const items = await Promise.all(
      accepted.map(async (file) => {
        const url = URL.createObjectURL(file);
        const image = await loadImage(url);
        return {
          id: `${file.name}-${file.lastModified}-${crypto.randomUUID()}`,
          file,
          url,
          width: image.naturalWidth,
          height: image.naturalHeight,
          rotation: 0,
        };
      }),
    );
    setImages((current) => [...current, ...items]);
    setNotice(`${items.length} 張圖片已加入，可拖曳調整順序。`);
  }, []);

  const handleDragEnd = ({ active, over }: DragEndEvent) => {
    if (!over || active.id === over.id) return;
    setImages((items) => {
      const oldIndex = items.findIndex((item) => item.id === active.id);
      const newIndex = items.findIndex((item) => item.id === over.id);
      return arrayMove(items, oldIndex, newIndex);
    });
  };

  const removeImage = (id: string) => {
    setImages((items) => {
      const target = items.find((item) => item.id === id);
      if (target) URL.revokeObjectURL(target.url);
      return items.filter((item) => item.id !== id);
    });
  };

  const clearAll = () => {
    images.forEach((item) => URL.revokeObjectURL(item.url));
    setImages([]);
    setNotice("");
  };

  const rotateImage = (id: string) => {
    setImages((items) => items.map((item) => (item.id === id ? { ...item, rotation: (item.rotation + 90) % 360 } : item)));
  };

  const moveImage = (from: number, to: number) => {
    if (to < 0 || to >= images.length) return;
    setImages((items) => arrayMove(items, from, to));
  };

  const createPdf = async () => {
    if (!images.length) return;
    setIsGenerating(true);
    setNotice("正在整理頁面並產生 PDF…");
    try {
      let pdf: jsPDF | null = null;
      for (let index = 0; index < images.length; index += 1) {
        const item = images[index];
        const source = await loadImage(item.url);
        const canvas = renderRotatedImage(source, item.rotation);
        const dataUrl = canvas.toDataURL("image/jpeg", 0.94);
        const format = pageSize === "match" ? [canvas.width * 0.75, canvas.height * 0.75] : pageSize;
        const pageOrientation = pageSize === "match" ? (canvas.width > canvas.height ? "landscape" : "portrait") : orientation;

        if (!pdf) {
          pdf = new jsPDF({ orientation: pageOrientation, unit: pageSize === "match" ? "pt" : "mm", format, compress: true });
        } else {
          pdf.addPage(format, pageOrientation);
        }

        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const pageMargin = pageSize === "match" ? 0 : margin;
        const maxWidth = pageWidth - pageMargin * 2;
        const maxHeight = pageHeight - pageMargin * 2;
        const scale = fitMode === "fill" && pageSize !== "match"
          ? Math.max(maxWidth / canvas.width, maxHeight / canvas.height)
          : Math.min(maxWidth / canvas.width, maxHeight / canvas.height);
        const imageWidth = canvas.width * scale;
        const imageHeight = canvas.height * scale;
        pdf.addImage(dataUrl, "JPEG", (pageWidth - imageWidth) / 2, (pageHeight - imageHeight) / 2, imageWidth, imageHeight, undefined, "FAST");
      }
      const safeName = (fileName.trim() || "圖片合輯").replace(/[\\/:*?"<>|]/g, "-");
      pdf?.save(`${safeName}.pdf`);
      setNotice(`完成！已將 ${images.length} 張圖片合併為 PDF。`);
    } catch (error) {
      console.error(error);
      setNotice("轉檔時發生問題，請確認圖片檔案後再試一次。");
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <main>
      <header className="topbar">
        <a className="brand" href="#top" aria-label="PictoPDF 首頁">
          <span className="brand-mark"><FileImage size={20} /></span>
          <span>Picto<span>PDF</span></span>
        </a>
        <div className="privacy-pill"><Lock size={14} /> 本機處理・圖片不上傳</div>
      </header>

      <section className="hero" id="top">
        <div className="eyebrow"><span /> JPG → PDF</div>
        <h1>圖片排好，<em>一鍵成冊。</em></h1>
        <p>把 JPG 拖進來、排成你要的頁序，幾秒內匯出乾淨的 PDF。</p>
      </section>

      <section className="workspace" aria-label="JPG 轉 PDF 工作區">
        <div className="main-column">
          <div
            className={`dropzone${isOver ? " is-over" : ""}${images.length ? " compact" : ""}`}
            onDragEnter={(event) => { event.preventDefault(); setIsOver(true); }}
            onDragOver={(event) => event.preventDefault()}
            onDragLeave={(event) => { if (event.currentTarget === event.target) setIsOver(false); }}
            onDrop={(event) => { event.preventDefault(); setIsOver(false); addFiles(event.dataTransfer.files); }}
          >
            <input ref={inputRef} type="file" accept="image/jpeg,.jpg,.jpeg" multiple hidden onChange={(event) => event.target.files && addFiles(event.target.files)} />
            <div className="upload-icon"><UploadCloud size={30} /></div>
            <div>
              <h2>{images.length ? "繼續加入圖片" : "將 JPG 拖放到這裡"}</h2>
              <p>{images.length ? "新圖片會接在目前頁序之後" : "支援一次選取多張 JPG / JPEG"}</p>
            </div>
            <button type="button" className="secondary-button" onClick={() => inputRef.current?.click()}><ImagePlus size={18} /> 選擇圖片</button>
          </div>

          {images.length > 0 && (
            <div className="queue-panel">
              <div className="queue-head">
                <div>
                  <span className="step-label">01 · 排列頁序</span>
                  <h2>{images.length} 張圖片</h2>
                </div>
                <button type="button" className="text-button" onClick={clearAll}><X size={16} /> 全部清除</button>
              </div>
              <p className="drag-tip"><GripVertical size={15} /> 按住卡片左上角拖曳，調整 PDF 頁面順序</p>
              <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={handleDragEnd}>
                <SortableContext items={images.map((item) => item.id)} strategy={rectSortingStrategy}>
                  <div className="image-grid">
                    {images.map((item, index) => (
                      <SortableCard key={item.id} item={item} index={index} total={images.length} onRemove={removeImage} onRotate={rotateImage} onMove={moveImage} />
                    ))}
                    <button type="button" className="add-card" onClick={() => inputRef.current?.click()}><ImagePlus size={24} /><span>加入更多</span></button>
                  </div>
                </SortableContext>
              </DndContext>
            </div>
          )}
        </div>

        <aside className="settings-panel">
          <div className="settings-heading">
            <span className="step-label">02 · PDF 設定</span>
            <h2>輸出樣式</h2>
          </div>

          <fieldset>
            <legend>頁面方向</legend>
            <div className="segmented two">
              <button type="button" className={orientation === "portrait" ? "active" : ""} onClick={() => setOrientation("portrait")} disabled={pageSize === "match"}>
                <span className="paper portrait" />直向{orientation === "portrait" && <Check size={14} />}
              </button>
              <button type="button" className={orientation === "landscape" ? "active" : ""} onClick={() => setOrientation("landscape")} disabled={pageSize === "match"}>
                <span className="paper landscape" />橫向{orientation === "landscape" && <Check size={14} />}
              </button>
            </div>
          </fieldset>

          <fieldset>
            <legend>頁面尺寸</legend>
            <div className="option-stack">
              {[
                ["a4", "A4", "210 × 297 mm"],
                ["letter", "美國信紙", "216 × 279 mm"],
                ["match", "符合圖片", "每頁依原圖尺寸"],
              ].map(([value, label, note]) => (
                <button key={value} type="button" className={pageSize === value ? "radio-row active" : "radio-row"} onClick={() => setPageSize(value as typeof pageSize)}>
                  <span className="radio-dot" /><span><b>{label}</b><small>{note}</small></span>
                </button>
              ))}
            </div>
          </fieldset>

          <fieldset>
            <legend>頁面邊距</legend>
            <div className="segmented three">
              {[[0, "無"], [3, "極窄"], [10, "標準"]].map(([value, label]) => (
                <button key={value} type="button" className={margin === value ? "active" : ""} onClick={() => setMargin(value as 0 | 3 | 10)} disabled={pageSize === "match"}>{label}</button>
              ))}
            </div>
          </fieldset>

          <fieldset>
            <legend>圖片填版方式</legend>
            <div className="segmented two">
              <button type="button" className={fitMode === "contain" ? "active" : ""} onClick={() => setFitMode("contain")} disabled={pageSize === "match"}>完整圖片（建議）</button>
              <button type="button" className={fitMode === "fill" ? "active" : ""} onClick={() => setFitMode("fill")} disabled={pageSize === "match"}>填滿（會裁切）</button>
            </div>
            <p className="option-note">預設保留整張圖片與所有文字；只有手動選擇「填滿」才會裁切邊緣。</p>
          </fieldset>

          <label className="filename-label">
            檔案名稱
            <div className="filename-input"><input value={fileName} onChange={(event) => setFileName(event.target.value)} /><span>.pdf</span></div>
          </label>

          <button className="primary-button" type="button" onClick={createPdf} disabled={!images.length || isGenerating}>
            {isGenerating ? <LoaderCircle className="spin" size={20} /> : <Download size={20} />}
            {isGenerating ? "正在建立 PDF" : images.length ? `下載 PDF · ${images.length} 頁` : "請先加入圖片"}
          </button>
          <p className="local-note"><Lock size={13} /> 轉檔全程在你的瀏覽器內完成</p>
        </aside>
      </section>

      <div className="status" role="status" aria-live="polite">{notice}</div>

      <section className="how-it-works">
        <span className="step-label">簡單三步驟</span>
        <div className="steps">
          <div><b>1</b><span><strong>加入 JPG</strong><small>拖放或一次多選圖片</small></span></div>
          <div><b>2</b><span><strong>拖曳排序</strong><small>確認每一頁的先後順序</small></span></div>
          <div><b>3</b><span><strong>下載 PDF</strong><small>套用設定並立即匯出</small></span></div>
        </div>
      </section>

      <footer><span>PictoPDF</span><p>快速、私密、免安裝的圖片轉 PDF 工具</p></footer>
    </main>
  );
}
