"use client";

import { useCallback, useRef, useState } from "react";

import { Button } from "@/components/ui/button";

interface GcashUploadDropzoneProps {
  readonly onFilesSelected: (files: readonly File[]) => void;
  readonly isProcessing: boolean;
}

export function GcashUploadDropzone({
  onFilesSelected,
  isProcessing,
}: GcashUploadDropzoneProps) {
  const [isDragging, setIsDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const selectFiles = useCallback((files: readonly File[]) => {
    const valid = files.filter((file) => file.name.split(".").pop()?.toLowerCase() === "xlsx");
    if (valid.length !== files.length) alert("Please upload Excel files (.xlsx)");
    if (valid.length) onFilesSelected(valid);
  }, [onFilesSelected]);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);
      const files = [...e.dataTransfer.files];
      if (files.length) selectFiles(files);
    },
    [selectFiles],
  );

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback(() => {
    setIsDragging(false);
  }, []);

  return (
    <div
      onDrop={handleDrop}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      className={`flex flex-col items-center justify-center rounded-xl border-2 border-dashed p-12 text-center transition-colors ${
        isDragging
          ? "border-primary bg-primary/5"
          : "border-muted-foreground/25 hover:border-muted-foreground/50"
      }`}
    >
      <input
        ref={inputRef}
        type="file"
        multiple
        accept=".xlsx"
        className="hidden"
        onChange={(e) => {
          const files = [...(e.target.files ?? [])];
          if (files.length) selectFiles(files);
        }}
      />

      <div className="mb-4 text-muted-foreground">
        <svg
          className="mx-auto size-10"
          fill="none"
          viewBox="0 0 24 24"
          stroke="currentColor"
          strokeWidth={1.5}
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5"
          />
        </svg>
      </div>

      <p className="mb-1 text-sm font-medium">
        Drop one or more GCash Statement Excel files here
      </p>
      <p className="mb-4 text-xs text-muted-foreground">
        .xlsx files only
      </p>

      <Button
        variant="outline"
        size="sm"
        disabled={isProcessing}
        onClick={() => inputRef.current?.click()}
      >
        {isProcessing ? "Processing..." : "Choose files"}
      </Button>
    </div>
  );
}
