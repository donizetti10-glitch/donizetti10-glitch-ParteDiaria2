import React, { useRef, useState } from "react";
import {
  Camera,
  ImagePlus,
  Trash2,
  AlertCircle,
  X,
  Maximize2,
  UploadCloud,
  CheckCircle2,
} from "lucide-react";
import { type ReportPhoto } from "@/lib/parte-diaria";

interface PhotoUploaderProps {
  photos: ReportPhoto[];
  onChange: (photos: ReportPhoto[]) => void;
  maxPhotos?: number;
}

/**
 * Resizes and compresses an image file to an optimized, lightweight JPEG dataURL.
 * Uses createImageBitmap if supported for speed and low memory, with robust fallbacks.
 */
async function compressImageSafely(file: File, maxDim = 720, quality = 0.65): Promise<string> {
  // Strategy 1: Native createImageBitmap (Fastest, low memory, handles EXIF orientation)
  if (typeof window !== "undefined" && "createImageBitmap" in window) {
    try {
      const bitmap = await createImageBitmap(file);
      let { width: w, height: h } = bitmap;
      if (w > maxDim || h > maxDim) {
        if (w > h) {
          h = Math.round((h * maxDim) / w);
          w = maxDim;
        } else {
          w = Math.round((w * maxDim) / h);
          h = maxDim;
        }
      }
      const canvas = document.createElement("canvas");
      canvas.width = Math.max(w, 1);
      canvas.height = Math.max(h, 1);
      const ctx = canvas.getContext("2d", { alpha: false });
      if (ctx) {
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, w, h);
        ctx.drawImage(bitmap, 0, 0, w, h);
        bitmap.close();
        const dataUrl = canvas.toDataURL("image/jpeg", quality);
        if (dataUrl && dataUrl.length > 200) {
          return dataUrl;
        }
      }
      bitmap.close();
    } catch (bitmapErr) {
      console.warn("createImageBitmap fallback to ObjectURL:", bitmapErr);
    }
  }

  // Strategy 2: URL.createObjectURL + HTMLImageElement
  try {
    const dataUrl = await new Promise<string>((resolve, reject) => {
      const objectUrl = URL.createObjectURL(file);
      const img = new Image();

      img.onload = () => {
        try {
          let w = img.naturalWidth || img.width;
          let h = img.naturalHeight || img.height;
          if (w > maxDim || h > maxDim) {
            if (w > h) {
              h = Math.round((h * maxDim) / w);
              w = maxDim;
            } else {
              w = Math.round((w * maxDim) / h);
              h = maxDim;
            }
          }
          const canvas = document.createElement("canvas");
          canvas.width = Math.max(w, 1);
          canvas.height = Math.max(h, 1);
          const ctx = canvas.getContext("2d", { alpha: false });
          if (!ctx) {
            URL.revokeObjectURL(objectUrl);
            reject(new Error("Canvas context não disponível"));
            return;
          }
          ctx.fillStyle = "#ffffff";
          ctx.fillRect(0, 0, w, h);
          ctx.drawImage(img, 0, 0, w, h);
          URL.revokeObjectURL(objectUrl);
          const resultUrl = canvas.toDataURL("image/jpeg", quality);
          resolve(resultUrl);
        } catch (canvasErr) {
          URL.revokeObjectURL(objectUrl);
          reject(canvasErr);
        }
      };

      img.onerror = () => {
        URL.revokeObjectURL(objectUrl);
        reject(new Error("Falha ao decodificar imagem via ObjectURL"));
      };

      img.src = objectUrl;
    });

    if (dataUrl && dataUrl.length > 200) {
      return dataUrl;
    }
  } catch (objErr) {
    console.warn("ObjectURL fallback to FileReader:", objErr);
  }

  // Strategy 3: FileReader base64 fallback
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const rawData = e.target?.result as string;
      if (!rawData) {
        reject(new Error("Não foi possível ler os dados da foto."));
        return;
      }
      const img = new Image();
      img.onload = () => {
        let w = img.width;
        let h = img.height;
        if (w > maxDim || h > maxDim) {
          if (w > h) {
            h = Math.round((h * maxDim) / w);
            w = maxDim;
          } else {
            w = Math.round((w * maxDim) / h);
            h = maxDim;
          }
        }
        const canvas = document.createElement("canvas");
        canvas.width = Math.max(w, 1);
        canvas.height = Math.max(h, 1);
        const ctx = canvas.getContext("2d");
        if (!ctx) {
          resolve(rawData);
          return;
        }
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
        const compressed = canvas.toDataURL("image/jpeg", quality);
        resolve(compressed);
      };
      img.onerror = () => {
        // As last resort, if it's already a data URL, return it
        if (rawData.startsWith("data:image")) {
          resolve(rawData);
        } else {
          reject(new Error("Formato de imagem não suportado pelo navegador."));
        }
      };
      img.src = rawData;
    };
    reader.onerror = () => reject(new Error("Erro ao ler arquivo da foto."));
    reader.readAsDataURL(file);
  });
}

export function PhotoUploader({ photos, onChange, maxPhotos = 12 }: PhotoUploaderProps) {
  const galleryInputRef = useRef<HTMLInputElement>(null);
  const cameraInputRef = useRef<HTMLInputElement>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingStatus, setProcessingStatus] = useState<string>("");
  const [warningMessage, setWarningMessage] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);
  const [previewPhoto, setPreviewPhoto] = useState<ReportPhoto | null>(null);
  const [isDragging, setIsDragging] = useState(false);

  const processFiles = async (files: File[]) => {
    if (!files || files.length === 0) return;

    setIsProcessing(true);
    setWarningMessage(null);
    setSuccessMessage(null);

    try {
      const currentCount = photos.length;
      const availableSlots = maxPhotos - currentCount;

      if (availableSlots <= 0) {
        setWarningMessage(`O limite de ${maxPhotos} fotos já foi atingido.`);
        setIsProcessing(false);
        return;
      }

      const filesToProcess = files.slice(0, availableSlots);

      if (files.length > availableSlots) {
        setWarningMessage(
          `Você selecionou ${files.length} fotos. Apenas ${availableSlots} ${
            availableSlots === 1 ? "foto foi adicionada" : "fotos foram adicionadas"
          } para respeitar o limite de ${maxPhotos} fotos.`,
        );
      }

      const newPhotos: ReportPhoto[] = [];
      const failedFiles: string[] = [];

      for (let i = 0; i < filesToProcess.length; i++) {
        const file = filesToProcess[i];
        setProcessingStatus(
          `Processando foto ${i + 1} de ${filesToProcess.length} (${file.name || "imagem"})...`,
        );

        try {
          const dataUrl = await compressImageSafely(file, 720, 0.65);
          newPhotos.push({
            id: `photo-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
            dataUrl,
            caption: "",
          });
        } catch (err) {
          console.error("Erro ao processar imagem:", file.name, err);
          failedFiles.push(file.name || `Foto #${i + 1}`);
        }
      }

      if (newPhotos.length > 0) {
        onChange([...photos, ...newPhotos]);
        setSuccessMessage(
          `${newPhotos.length} ${newPhotos.length === 1 ? "foto adicionada" : "fotos adicionadas"} com sucesso!`,
        );
        setTimeout(() => setSuccessMessage(null), 4000);
      }

      if (failedFiles.length > 0) {
        setWarningMessage(
          `Não foi possível carregar ${failedFiles.length} foto(s): ${failedFiles.join(", ")}. Tente fotos nos formatos JPG, PNG ou WEBP.`,
        );
      }
    } catch (err) {
      console.error("Erro ao adicionar fotos:", err);
      setWarningMessage("Ocorreu um erro ao carregar as fotos. Tente novamente.");
    } finally {
      setIsProcessing(false);
      setProcessingStatus("");
      if (galleryInputRef.current) galleryInputRef.current.value = "";
      if (cameraInputRef.current) cameraInputRef.current.value = "";
    }
  };

  const handleFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    await processFiles(files);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    const droppedFiles = Array.from(e.dataTransfer.files || []).filter((f) =>
      f.type.startsWith("image/"),
    );
    if (droppedFiles.length > 0) {
      await processFiles(droppedFiles);
    }
  };

  const handleRemovePhoto = (id: string) => {
    onChange(photos.filter((p) => p.id !== id));
    setWarningMessage(null);
  };

  const handleUpdateCaption = (id: string, caption: string) => {
    onChange(photos.map((p) => (p.id === id ? { ...p, caption } : p)));
  };

  return (
    <div
      className={`bg-card border rounded-xl p-4 sm:p-5 space-y-4 shadow-sm transition-colors ${
        isDragging ? "border-primary bg-primary/5" : "border-border"
      }`}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-border/60">
        <div>
          <div className="flex items-center gap-2">
            <Camera className="w-5 h-5 text-primary shrink-0" />
            <h3 className="font-semibold text-foreground text-sm sm:text-base">
              Fotos para o Relatório PDF (Opcional)
            </h3>
            <span className="text-xs px-2 py-0.5 rounded-full font-semibold bg-primary/10 text-primary border border-primary/20">
              {photos.length} de {maxPhotos}
            </span>
          </div>
          <p className="text-xs text-muted-foreground mt-1">
            Anexe de 1 a {maxPhotos} fotos. Elas são compactadas automaticamente e diagramadas sem
            distorção no PDF.
          </p>
        </div>

        {/* Action buttons (Camera + Gallery) */}
        <div className="flex flex-wrap items-center gap-2">
          {/* Hidden inputs */}
          <input
            type="file"
            ref={galleryInputRef}
            onChange={handleFileSelect}
            accept="image/*,image/jpeg,image/png,image/webp,image/heic,image/heif"
            multiple
            className="hidden"
            disabled={photos.length >= maxPhotos || isProcessing}
          />
          <input
            type="file"
            ref={cameraInputRef}
            onChange={handleFileSelect}
            accept="image/*"
            capture="environment"
            className="hidden"
            disabled={photos.length >= maxPhotos || isProcessing}
          />

          {/* Camera Button */}
          <button
            type="button"
            onClick={() => cameraInputRef.current?.click()}
            disabled={photos.length >= maxPhotos || isProcessing}
            className={`btn inline-flex items-center gap-1.5 text-xs sm:text-sm font-semibold px-3 py-2 rounded-lg transition-all ${
              photos.length >= maxPhotos || isProcessing
                ? "bg-muted text-muted-foreground cursor-not-allowed"
                : "bg-secondary text-secondary-foreground hover:bg-secondary/80 border border-border/60 shadow-xs"
            }`}
            title="Tirar foto direto com a câmera do celular"
          >
            <Camera className="w-4 h-4 text-primary" />
            <span>Câmera</span>
          </button>

          {/* Gallery / File Button */}
          <button
            type="button"
            onClick={() => galleryInputRef.current?.click()}
            disabled={photos.length >= maxPhotos || isProcessing}
            className={`btn inline-flex items-center gap-1.5 text-xs sm:text-sm font-semibold px-3.5 py-2 rounded-lg transition-all ${
              photos.length >= maxPhotos || isProcessing
                ? "bg-muted text-muted-foreground cursor-not-allowed"
                : "bg-primary text-primary-foreground hover:bg-primary/90 shadow-xs"
            }`}
          >
            {isProcessing ? (
              <>
                <div className="w-4 h-4 border-2 border-primary-foreground border-t-transparent rounded-full animate-spin" />
                <span>Processando...</span>
              </>
            ) : (
              <>
                <ImagePlus className="w-4 h-4" />
                <span>
                  {photos.length === 0
                    ? "Adicionar Fotos"
                    : photos.length >= maxPhotos
                      ? "Limite Atingido"
                      : "Mais Fotos"}
                </span>
              </>
            )}
          </button>
        </div>
      </div>

      {/* Processing Status Banner */}
      {isProcessing && processingStatus && (
        <div className="flex items-center gap-2.5 text-xs bg-primary/10 text-primary p-3 rounded-lg border border-primary/20 animate-pulse">
          <div className="w-4 h-4 border-2 border-primary border-t-transparent rounded-full animate-spin shrink-0" />
          <span className="font-medium">{processingStatus}</span>
        </div>
      )}

      {/* Success Notification */}
      {successMessage && (
        <div className="flex items-center gap-2 text-xs bg-emerald-500/10 text-emerald-700 dark:text-emerald-300 p-3 rounded-lg border border-emerald-500/20 animate-in fade-in">
          <CheckCircle2 className="w-4 h-4 shrink-0 text-emerald-600 dark:text-emerald-400" />
          <span className="font-medium">{successMessage}</span>
        </div>
      )}

      {/* Warning / Error Notification */}
      {warningMessage && (
        <div className="flex items-start gap-2 text-xs bg-amber-500/10 text-amber-700 dark:text-amber-400 p-3 rounded-lg border border-amber-500/20">
          <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
          <span>{warningMessage}</span>
        </div>
      )}

      {/* Photo List or Empty State */}
      {photos.length === 0 ? (
        <div
          onClick={() => galleryInputRef.current?.click()}
          className="text-center py-8 px-4 border-2 border-dashed border-border/80 rounded-xl bg-muted/20 hover:bg-muted/40 cursor-pointer transition-colors"
        >
          <UploadCloud className="w-9 h-9 text-primary mx-auto mb-2 opacity-80" />
          <p className="text-xs sm:text-sm text-foreground font-semibold">
            Clique aqui ou arraste fotos para anexar ao relatório
          </p>
          <p className="text-[11px] text-muted-foreground mt-1 max-w-md mx-auto">
            Você pode tirar fotos pelo celular ou escolher da galeria (formatos JPG, PNG, WEBP, até{" "}
            {maxPhotos} fotos).
          </p>
        </div>
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3.5">
          {photos.map((photo, index) => (
            <div
              key={photo.id}
              className="group relative bg-background border border-border rounded-lg overflow-hidden flex flex-col shadow-xs transition-all hover:border-primary/50"
            >
              <div className="relative aspect-4/3 bg-muted/40 overflow-hidden flex items-center justify-center p-1">
                <img
                  src={photo.dataUrl}
                  alt={`Foto ${index + 1}`}
                  className="w-full h-full object-contain rounded cursor-pointer"
                  onClick={() => setPreviewPhoto(photo)}
                  referrerPolicy="no-referrer"
                />
                <span className="absolute top-2 left-2 bg-background/90 backdrop-blur-xs text-foreground font-bold text-[10px] px-2 py-0.5 rounded shadow-xs border border-border/60">
                  Foto {(index + 1).toString().padStart(2, "0")}
                </span>

                {/* View Full Size */}
                <button
                  type="button"
                  onClick={() => setPreviewPhoto(photo)}
                  className="absolute bottom-2 left-2 bg-background/80 text-foreground hover:bg-background p-1.5 rounded-md shadow-xs transition-transform opacity-80 hover:opacity-100"
                  title="Visualizar em tamanho maior"
                >
                  <Maximize2 className="w-3.5 h-3.5" />
                </button>

                {/* Delete Photo */}
                <button
                  type="button"
                  onClick={() => handleRemovePhoto(photo.id)}
                  className="absolute top-2 right-2 bg-destructive text-destructive-foreground hover:bg-destructive/90 p-1.5 rounded-full shadow-md transition-transform hover:scale-110"
                  title="Remover esta foto"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                </button>
              </div>

              <div className="p-2.5 flex-1 flex flex-col justify-between bg-card">
                <input
                  type="text"
                  value={photo.caption || ""}
                  onChange={(e) => handleUpdateCaption(photo.id, e.target.value)}
                  placeholder={`Legenda da Foto ${(index + 1).toString().padStart(2, "0")} (opcional)...`}
                  className="w-full text-xs px-2.5 py-1.5 bg-muted/30 border border-input rounded focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background transition-colors"
                />
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Lightbox / Preview Modal */}
      {previewPhoto && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 p-4 animate-in fade-in"
          onClick={() => setPreviewPhoto(null)}
        >
          <div
            className="relative max-w-3xl w-full bg-card border border-border rounded-xl p-4 shadow-2xl space-y-3"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between pb-2 border-b border-border">
              <h4 className="font-semibold text-sm sm:text-base text-foreground">
                Visualização da Foto
              </h4>
              <button
                type="button"
                onClick={() => setPreviewPhoto(null)}
                className="p-1.5 rounded-lg text-muted-foreground hover:text-foreground hover:bg-muted"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="max-h-[70vh] flex items-center justify-center overflow-hidden rounded-lg bg-black/5">
              <img
                src={previewPhoto.dataUrl}
                alt="Visualização"
                className="max-h-[65vh] w-auto max-w-full object-contain rounded"
                referrerPolicy="no-referrer"
              />
            </div>

            {previewPhoto.caption && (
              <p className="text-xs text-muted-foreground italic text-center">
                "{previewPhoto.caption}"
              </p>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
