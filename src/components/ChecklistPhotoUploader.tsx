import React, { useRef, useState } from "react";
import {
  Camera,
  ImagePlus,
  Trash2,
  X,
  Maximize2,
  CheckCircle2,
  AlertTriangle,
  Tag,
  Loader2,
} from "lucide-react";
import type { ChecklistPhoto, ChecklistItem } from "@/lib/checklist";

interface ChecklistPhotoUploaderProps {
  photos: ChecklistPhoto[];
  onChange: (photos: ChecklistPhoto[]) => void;
  availableItems?: ChecklistItem[];
  maxPhotos?: number;
}

/**
 * Comprime e redimensiona fotos da câmera ou galeria do celular de forma segura e rápida.
 * Utiliza createImageBitmap quando disponível (lida automaticamente com orientação EXIF e baixo uso de RAM).
 */
async function compressImageSafely(file: File, maxDim = 800, quality = 0.7): Promise<string> {
  // Estratégia 1: createImageBitmap nativo
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

  // Estratégia 2: URL.createObjectURL + HTMLImageElement
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
        reject(new Error("Falha ao carregar imagem"));
      };

      img.src = objectUrl;
    });

    if (dataUrl && dataUrl.length > 200) {
      return dataUrl;
    }
  } catch (objErr) {
    console.warn("ObjectURL fallback to FileReader:", objErr);
  }

  // Estratégia 3: FileReader base64
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const rawData = e.target?.result as string;
      if (!rawData) {
        reject(new Error("Não foi possível ler os dados da foto"));
        return;
      }
      resolve(rawData);
    };
    reader.onerror = () => reject(new Error("Erro ao ler foto do dispositivo"));
    reader.readAsDataURL(file);
  });
}

export function ChecklistPhotoUploader({
  photos,
  onChange,
  availableItems = [],
  maxPhotos = 15,
}: ChecklistPhotoUploaderProps) {
  const cameraInputRef = useRef<HTMLInputElement>(null);
  const galleryInputRef = useRef<HTMLInputElement>(null);

  const [isProcessing, setIsProcessing] = useState(false);
  const [processingStatus, setProcessingStatus] = useState<string>("");
  const [warningMessage, setWarningMessage] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);
  const [previewPhoto, setPreviewPhoto] = useState<ChecklistPhoto | null>(null);
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
        setWarningMessage(`Limite de ${maxPhotos} fotos atingido.`);
        setIsProcessing(false);
        return;
      }

      const filesToProcess = files.slice(0, availableSlots);

      if (files.length > availableSlots) {
        setWarningMessage(
          `Foram selecionadas ${files.length} fotos. Apenas ${availableSlots} ${
            availableSlots === 1 ? "foto foi anexada" : "fotos foram anexadas"
          } pelo limite configurado de ${maxPhotos}.`,
        );
      }

      const newPhotos: ChecklistPhoto[] = [];
      const failedFiles: string[] = [];

      for (let i = 0; i < filesToProcess.length; i++) {
        const file = filesToProcess[i];
        setProcessingStatus(
          `Otimizando foto ${i + 1} de ${filesToProcess.length} da câmera nativa...`,
        );

        try {
          const dataUrl = await compressImageSafely(file, 800, 0.7);
          newPhotos.push({
            id: `photo-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
            dataUrl,
            caption: "",
            timestamp: Date.now(),
          });
        } catch (err) {
          console.error("Erro ao processar imagem:", file.name, err);
          failedFiles.push(file.name || `Foto #${i + 1}`);
        }
      }

      if (newPhotos.length > 0) {
        onChange([...photos, ...newPhotos]);
        setSuccessMessage(
          `${newPhotos.length} ${newPhotos.length === 1 ? "foto registrada" : "fotos registradas"} com sucesso!`,
        );
        setTimeout(() => setSuccessMessage(null), 4000);
      }

      if (failedFiles.length > 0) {
        setWarningMessage(`Não foi possível carregar ${failedFiles.length} foto(s).`);
      }
    } catch (err) {
      console.error("Erro ao adicionar fotos:", err);
      setWarningMessage("Erro ao capturar ou carregar foto. Tente novamente.");
    } finally {
      setIsProcessing(false);
      setProcessingStatus("");
      if (cameraInputRef.current) cameraInputRef.current.value = "";
      if (galleryInputRef.current) galleryInputRef.current.value = "";
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
    if (previewPhoto?.id === id) setPreviewPhoto(null);
  };

  const handleUpdateCaption = (id: string, caption: string) => {
    onChange(photos.map((p) => (p.id === id ? { ...p, caption } : p)));
  };

  const handleUpdateAssociatedItem = (photoId: string, itemId: string) => {
    const matched = availableItems.find((it) => it.id === itemId);
    onChange(
      photos.map((p) =>
        p.id === photoId
          ? {
              ...p,
              itemId: itemId || undefined,
              itemName: matched ? matched.name : undefined,
            }
          : p,
      ),
    );
  };

  return (
    <div
      id="checklist-photo-uploader"
      className={`bg-card border rounded-xl p-4 sm:p-5 space-y-4 shadow-sm transition-all ${
        isDragging ? "border-primary bg-primary/5 ring-2 ring-primary/20" : "border-border"
      }`}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {/* Inputs invisíveis */}
      {/* 1. Câmera Nativa do Celular (capture="environment") */}
      <input
        type="file"
        ref={cameraInputRef}
        onChange={handleFileSelect}
        accept="image/*"
        capture="environment"
        className="hidden"
        disabled={photos.length >= maxPhotos || isProcessing}
        id="camera-input-native"
      />

      {/* 2. Galeria / Arquivos */}
      <input
        type="file"
        ref={galleryInputRef}
        onChange={handleFileSelect}
        accept="image/*,image/jpeg,image/png,image/webp,image/heic,image/heif"
        multiple
        className="hidden"
        disabled={photos.length >= maxPhotos || isProcessing}
        id="gallery-input-files"
      />

      {/* Header com Ações */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-border/70">
        <div>
          <div className="flex items-center gap-2">
            <div className="p-2 rounded-lg bg-emerald-500/10 text-emerald-600 dark:text-emerald-400">
              <Camera className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="font-semibold text-foreground text-sm sm:text-base">
                  Fotos e Evidências da Inspeção
                </h3>
                <span className="text-xs px-2 py-0.5 rounded-full font-semibold bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/20">
                  {photos.length} de {maxPhotos}
                </span>
              </div>
              <p className="text-xs text-muted-foreground mt-0.5">
                Utilize a <strong>câmera nativa do celular</strong> para registrar anomalias,
                vazamentos, estado de peças ou anexe fotos da galeria.
              </p>
            </div>
          </div>
        </div>

        {/* Botões de Ação */}
        <div className="flex flex-wrap items-center gap-2">
          {/* Botão Câmera Nativa */}
          <button
            type="button"
            id="checklist-take-photo-btn"
            onClick={() => cameraInputRef.current?.click()}
            disabled={photos.length >= maxPhotos || isProcessing}
            className={`inline-flex items-center justify-center gap-2 text-xs sm:text-sm font-semibold px-4 py-2.5 rounded-lg transition-all shadow-xs ${
              photos.length >= maxPhotos || isProcessing
                ? "bg-muted text-muted-foreground cursor-not-allowed"
                : "bg-emerald-600 hover:bg-emerald-700 text-white active:scale-98"
            }`}
            title="Abrir a câmera nativa do celular para tirar uma foto agora"
          >
            <Camera className="w-4 h-4" />
            <span>Tirar Foto (Câmera)</span>
          </button>

          {/* Botão Galeria / Arquivos */}
          <button
            type="button"
            id="checklist-attach-photo-btn"
            onClick={() => galleryInputRef.current?.click()}
            disabled={photos.length >= maxPhotos || isProcessing}
            className={`inline-flex items-center justify-center gap-2 text-xs sm:text-sm font-semibold px-3.5 py-2.5 rounded-lg transition-all border shadow-xs ${
              photos.length >= maxPhotos || isProcessing
                ? "bg-muted text-muted-foreground border-border cursor-not-allowed"
                : "bg-card text-foreground hover:bg-muted border-border/80"
            }`}
            title="Selecionar fotos existentes na galeria do celular ou computador"
          >
            <ImagePlus className="w-4 h-4 text-primary" />
            <span>Anexar da Galeria</span>
          </button>
        </div>
      </div>

      {/* Avisos e Status */}
      {isProcessing && (
        <div className="flex items-center gap-2.5 p-3 rounded-lg bg-primary/10 border border-primary/20 text-xs font-medium text-foreground animate-pulse">
          <Loader2 className="w-4 h-4 text-primary animate-spin" />
          <span>{processingStatus || "Processando foto..."}</span>
        </div>
      )}

      {successMessage && (
        <div className="flex items-center gap-2 p-2.5 rounded-lg bg-emerald-500/10 border border-emerald-500/20 text-xs text-emerald-700 dark:text-emerald-300 font-medium">
          <CheckCircle2 className="w-4 h-4 shrink-0 text-emerald-600 dark:text-emerald-400" />
          <span>{successMessage}</span>
        </div>
      )}

      {warningMessage && (
        <div className="flex items-center gap-2 p-2.5 rounded-lg bg-amber-500/10 border border-amber-500/20 text-xs text-amber-700 dark:text-amber-300 font-medium">
          <AlertTriangle className="w-4 h-4 shrink-0 text-amber-600 dark:text-amber-400" />
          <span>{warningMessage}</span>
        </div>
      )}

      {/* Empty State ou Grid de Fotos */}
      {photos.length === 0 ? (
        <div
          onClick={() => cameraInputRef.current?.click()}
          className="border-2 border-dashed border-border/80 hover:border-emerald-500/50 rounded-xl p-6 text-center cursor-pointer transition-colors bg-muted/10 hover:bg-emerald-500/5"
        >
          <div className="w-12 h-12 rounded-full bg-emerald-500/10 text-emerald-600 flex items-center justify-center mx-auto mb-2.5">
            <Camera className="w-6 h-6" />
          </div>
          <div className="font-semibold text-xs sm:text-sm text-foreground">
            Nenhuma foto anexada neste checklist
          </div>
          <p className="text-xs text-muted-foreground mt-1 max-w-md mx-auto">
            Toque aqui ou no botão acima para <strong>abrir a câmera do celular</strong> e
            fotografar o equipamento ou componentes avariados.
          </p>
          <div className="mt-3 flex items-center justify-center gap-2 text-[11px] text-muted-foreground">
            <span className="inline-flex items-center gap-1 bg-muted px-2 py-0.5 rounded">
              📷 Câmera Traseira Direta
            </span>
            <span className="inline-flex items-center gap-1 bg-muted px-2 py-0.5 rounded">
              🖼️ JPG / PNG / WEBP
            </span>
          </div>
        </div>
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3.5">
          {photos.map((photo, index) => (
            <div
              key={photo.id}
              className="bg-muted/20 border border-border/80 rounded-xl overflow-hidden shadow-xs hover:shadow-sm transition-all flex flex-col"
            >
              {/* Imagem + Overlay */}
              <div className="relative aspect-4/3 bg-black/5 overflow-hidden group">
                <img
                  src={photo.dataUrl}
                  alt={photo.caption || `Foto de inspeção ${index + 1}`}
                  className="w-full h-full object-cover transition-transform group-hover:scale-102"
                />

                {/* Badge com número */}
                <div className="absolute top-2 left-2 px-2 py-0.5 rounded bg-black/70 text-white font-mono text-[10px] font-semibold backdrop-blur-xs">
                  #{index + 1}
                </div>

                {/* Botões de Ação sobre a imagem */}
                <div className="absolute top-2 right-2 flex items-center gap-1.5 opacity-90 group-hover:opacity-100 transition-opacity">
                  <button
                    type="button"
                    onClick={() => setPreviewPhoto(photo)}
                    className="p-1.5 rounded-lg bg-black/60 hover:bg-black/80 text-white backdrop-blur-xs transition-colors"
                    title="Ampliar foto em tela cheia"
                  >
                    <Maximize2 className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onClick={() => handleRemovePhoto(photo.id)}
                    className="p-1.5 rounded-lg bg-red-600/80 hover:bg-red-600 text-white backdrop-blur-xs transition-colors"
                    title="Remover esta foto"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              </div>

              {/* Controles da Foto: Legenda e Associação com Item */}
              <div className="p-3 space-y-2 flex-1 flex flex-col justify-between">
                <div className="space-y-1.5">
                  <div className="flex items-center justify-between gap-1 text-[11px] font-medium text-muted-foreground">
                    <span className="flex items-center gap-1">
                      <Tag className="w-3 h-3 text-primary" />
                      Associar ao Item:
                    </span>
                  </div>
                  <select
                    value={photo.itemId || ""}
                    onChange={(e) => handleUpdateAssociatedItem(photo.id, e.target.value)}
                    className="w-full text-xs p-1.5 bg-background border border-input rounded-md focus:outline-none focus:ring-1 focus:ring-primary truncate"
                  >
                    <option value="">Geral / Equipamento Completo</option>
                    {availableItems.map((item) => (
                      <option key={item.id} value={item.id}>
                        {item.category}: {item.name}
                      </option>
                    ))}
                  </select>
                </div>

                <div className="space-y-1">
                  <input
                    type="text"
                    value={photo.caption || ""}
                    onChange={(e) => handleUpdateCaption(photo.id, e.target.value)}
                    placeholder="Descrição da avaria ou componente..."
                    className="w-full text-xs p-1.5 bg-background border border-input rounded-md focus:outline-none focus:ring-1 focus:ring-primary"
                  />
                </div>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Modal de Pré-visualização em Tela Cheia */}
      {previewPhoto && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/85 p-3 sm:p-6 animate-in fade-in"
          onClick={() => setPreviewPhoto(null)}
        >
          <div
            className="relative max-w-4xl w-full bg-card rounded-2xl overflow-hidden shadow-2xl border border-border flex flex-col max-h-[92vh]"
            onClick={(e) => e.stopPropagation()}
          >
            {/* Modal Header */}
            <div className="p-3.5 sm:p-4 border-b border-border flex items-center justify-between bg-muted/30">
              <div className="flex items-center gap-2">
                <Camera className="w-4 h-4 text-emerald-600" />
                <h4 className="font-semibold text-xs sm:text-sm text-foreground">
                  Foto de Inspeção{" "}
                  {previewPhoto.itemName ? `• ${previewPhoto.itemName}` : "• Geral"}
                </h4>
              </div>
              <button
                type="button"
                onClick={() => setPreviewPhoto(null)}
                className="p-1.5 rounded-lg hover:bg-muted text-muted-foreground hover:text-foreground"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Image */}
            <div className="flex-1 overflow-auto bg-black flex items-center justify-center p-2">
              <img
                src={previewPhoto.dataUrl}
                alt={previewPhoto.caption || "Foto ampliada"}
                className="max-h-[70vh] w-auto object-contain rounded"
              />
            </div>

            {/* Modal Footer */}
            <div className="p-3.5 bg-card border-t border-border flex items-center justify-between gap-3">
              <div className="text-xs text-muted-foreground truncate">
                {previewPhoto.caption ? (
                  <span className="text-foreground font-medium">"{previewPhoto.caption}"</span>
                ) : (
                  <span>Sem legenda registrada</span>
                )}
              </div>
              <button
                type="button"
                onClick={() => handleRemovePhoto(previewPhoto.id)}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-destructive/10 text-destructive hover:bg-destructive/20 transition-colors shrink-0"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Excluir Foto</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
