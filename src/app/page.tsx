"use client";

import {
  AlignmentType,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  TextRun,
} from "docx";
import {
  type ChangeEvent,
  type DragEvent,
  useMemo,
  useRef,
  useState,
} from "react";

const fontChoices = ["Aptos", "Calibri", "Georgia", "Inter"] as const;
const pageSizeChoices = [
  { value: "a4", label: "A4" },
  { value: "letter", label: "Letter" },
] as const;

type FontChoice = (typeof fontChoices)[number];
type PageSize = (typeof pageSizeChoices)[number]["value"];

const pageSizes: Record<PageSize, { readonly width: number; readonly height: number }> = {
  a4: { width: 11906, height: 16838 },
  letter: { width: 12240, height: 15840 },
};

const defaultText = `Paste your text here or upload a .txt file.

This converter creates a real .docx file in your browser.
No upload. No server. Just clean Word output.`;

const cleanDownloadName = (name: string): string => {
  const baseName = name.trim().replace(/\.txt$/i, "") || "converted-document";
  return baseName.replace(/[^a-z0-9-_ ]/gi, "").replace(/\s+/g, "-").toLowerCase();
};

const countWords = (value: string): number => {
  const words = value.trim().match(/\S+/g);
  return words === null ? 0 : words.length;
};

const isFontChoice = (value: string): value is FontChoice =>
  fontChoices.some((font) => font === value);

const isPageSize = (value: string): value is PageSize =>
  pageSizeChoices.some((size) => size.value === value);

const saveBlob = (blob: Blob, fileName: string): void => {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
};

export default function Home() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [text, setText] = useState(defaultText);
  const [fileName, setFileName] = useState("converted-document");
  const [documentTitle, setDocumentTitle] = useState("Converted Document");
  const [fontFamily, setFontFamily] = useState<FontChoice>("Aptos");
  const [fontSize, setFontSize] = useState(12);
  const [pageSize, setPageSize] = useState<PageSize>("a4");
  const [includeTitle, setIncludeTitle] = useState(true);
  const [isDragging, setIsDragging] = useState(false);
  const [status, setStatus] = useState("Ready to convert");
  const [isConverting, setIsConverting] = useState(false);

  const stats = useMemo(
    () => ({
      characters: text.length,
      words: countWords(text),
      lines: text.length === 0 ? 0 : text.replace(/\r\n/g, "\n").split("\n").length,
    }),
    [text],
  );

  const hasContent = text.trim().length > 0;

  const loadFile = async (file: File): Promise<void> => {
    if (!file.name.toLowerCase().endsWith(".txt")) {
      setStatus("Please choose a .txt file.");
      return;
    }

    const fileText = await file.text();
    setText(fileText);
    setFileName(cleanDownloadName(file.name));
    setDocumentTitle(file.name.replace(/\.txt$/i, "") || "Converted Document");
    setStatus(`${file.name} loaded successfully.`);
  };

  const handleFileChange = async (event: ChangeEvent<HTMLInputElement>): Promise<void> => {
    const files = event.target.files;
    if (files === null || files.length === 0) {
      return;
    }

    await loadFile(files[0]);
    event.target.value = "";
  };

  const handleDrop = async (event: DragEvent<HTMLLabelElement>): Promise<void> => {
    event.preventDefault();
    setIsDragging(false);

    const file = event.dataTransfer.files[0];
    if (file instanceof File) {
      await loadFile(file);
    }
  };

  const createDocument = (): Document => {
    const normalizedLines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    const bodyParagraphs = normalizedLines.map(
      (line) =>
        new Paragraph({
          children: [
            new TextRun({
              text: line.length === 0 ? " " : line,
              font: fontFamily,
              size: fontSize * 2,
            }),
          ],
          spacing: {
            after: line.trim().length === 0 ? 80 : 160,
          },
        }),
    );

    const titleParagraphs = includeTitle
      ? [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            heading: HeadingLevel.TITLE,
            spacing: { after: 360 },
            children: [
              new TextRun({
                text: documentTitle.trim() || "Converted Document",
                font: fontFamily,
                bold: true,
                size: 34,
              }),
            ],
          }),
        ]
      : [];

    return new Document({
      creator: "TXT to Word Converter",
      title: documentTitle,
      description: "Generated from a plain text file.",
      sections: [
        {
          properties: {
            page: {
              size: {
                width: pageSizes[pageSize].width,
                height: pageSizes[pageSize].height,
              },
              margin: {
                top: 1440,
                right: 1440,
                bottom: 1440,
                left: 1440,
              },
            },
          },
          children: [...titleParagraphs, ...bodyParagraphs],
        },
      ],
    });
  };

  const handleConvert = async (): Promise<void> => {
    if (!hasContent) {
      setStatus("Add text before converting.");
      return;
    }

    setIsConverting(true);
    setStatus("Creating your Word document...");

    try {
      const document = createDocument();
      const blob = await Packer.toBlob(document);
      saveBlob(blob, `${cleanDownloadName(fileName)}.docx`);
      setStatus("DOCX downloaded successfully.");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unexpected conversion error.";
      setStatus(message);
    } finally {
      setIsConverting(false);
    }
  };

  return (
    <main className="min-h-screen overflow-hidden bg-[#0b1020] px-4 py-6 text-slate-950 sm:px-6 lg:px-8">
      <div className="pointer-events-none fixed inset-0 -z-10">
        <div className="absolute left-[-10%] top-[-15%] h-96 w-96 rounded-full bg-fuchsia-500/30 blur-3xl" />
        <div className="absolute right-[-8%] top-[12%] h-[28rem] w-[28rem] rounded-full bg-cyan-400/25 blur-3xl" />
        <div className="absolute bottom-[-18%] left-[30%] h-[30rem] w-[30rem] rounded-full bg-violet-500/20 blur-3xl" />
      </div>

      <section className="mx-auto flex min-h-[calc(100vh-3rem)] w-full max-w-7xl flex-col rounded-[2rem] border border-white/15 bg-white/90 p-4 shadow-2xl shadow-black/30 backdrop-blur-2xl sm:p-6 lg:p-8">
        <header className="grid gap-8 py-5 lg:grid-cols-[1.05fr_0.95fr] lg:items-end">
          <div>
            <div className="mb-5 inline-flex items-center gap-2 rounded-full border border-slate-200 bg-white px-4 py-2 text-sm font-semibold text-slate-600 shadow-sm">
              <span className="h-2 w-2 rounded-full bg-emerald-400 shadow-[0_0_18px_rgba(52,211,153,0.9)]" />
              Private browser conversion
            </div>
            <h1 className="max-w-4xl text-5xl font-semibold tracking-[-0.05em] text-slate-950 sm:text-6xl lg:text-7xl">
              Convert TXT into polished Word files.
            </h1>
            <p className="mt-5 max-w-2xl text-lg leading-8 text-slate-600">
              Upload a text file or paste content, tune the document style, and
              export a real Word-compatible DOCX instantly.
            </p>
          </div>

          <div className="grid grid-cols-3 gap-3 rounded-3xl border border-slate-200 bg-slate-950 p-3 text-white shadow-xl shadow-slate-950/20">
            {[
              { label: "Words", value: stats.words.toLocaleString() },
              { label: "Lines", value: stats.lines.toLocaleString() },
              { label: "Chars", value: stats.characters.toLocaleString() },
            ].map((item) => (
              <div className="rounded-2xl bg-white/10 p-4" key={item.label}>
                <div className="text-2xl font-semibold tracking-tight">{item.value}</div>
                <div className="mt-1 text-xs font-medium uppercase tracking-[0.2em] text-slate-400">
                  {item.label}
                </div>
              </div>
            ))}
          </div>
        </header>

        <div className="grid flex-1 gap-5 lg:grid-cols-[0.92fr_1.08fr]">
          <aside className="space-y-5">
            <label
              className={`group flex min-h-72 cursor-pointer flex-col items-center justify-center rounded-[1.75rem] border border-dashed p-8 text-center transition ${
                isDragging
                  ? "border-cyan-400 bg-cyan-50 shadow-2xl shadow-cyan-300/30"
                  : "border-slate-300 bg-slate-50 hover:border-slate-400 hover:bg-white"
              }`}
              onDragEnter={() => setIsDragging(true)}
              onDragLeave={() => setIsDragging(false)}
              onDragOver={(event) => event.preventDefault()}
              onDrop={handleDrop}
            >
              <input
                accept=".txt,text/plain"
                className="sr-only"
                onChange={handleFileChange}
                ref={fileInputRef}
                type="file"
              />
              <span className="mb-6 grid h-20 w-20 place-items-center rounded-3xl bg-slate-950 text-3xl text-white shadow-2xl shadow-slate-950/30 transition group-hover:scale-105">
                TXT
              </span>
              <span className="text-2xl font-semibold tracking-tight text-slate-950">
                Drop your .txt file here
              </span>
              <span className="mt-3 max-w-xs text-sm leading-6 text-slate-500">
                Or click to browse. Your file is read locally and converted in
                this browser tab.
              </span>
            </label>

            <div className="rounded-[1.75rem] border border-slate-200 bg-white p-5 shadow-sm">
              <div className="mb-4 flex items-center justify-between gap-4">
                <h2 className="text-lg font-semibold tracking-tight text-slate-950">
                  Document settings
                </h2>
                <span className="rounded-full bg-emerald-50 px-3 py-1 text-xs font-semibold text-emerald-700">
                  DOCX
                </span>
              </div>

              <div className="grid gap-4">
                <label className="grid gap-2 text-sm font-medium text-slate-700">
                  Document title
                  <input
                    className="rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-slate-950 outline-none transition focus:border-cyan-400 focus:bg-white focus:ring-4 focus:ring-cyan-100"
                    onChange={(event) => setDocumentTitle(event.target.value)}
                    value={documentTitle}
                  />
                </label>

                <div className="grid gap-4 sm:grid-cols-2">
                  <label className="grid gap-2 text-sm font-medium text-slate-700">
                    Font
                    <select
                      className="rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-slate-950 outline-none transition focus:border-cyan-400 focus:bg-white focus:ring-4 focus:ring-cyan-100"
                      onChange={(event) => {
                        if (isFontChoice(event.target.value)) {
                          setFontFamily(event.target.value);
                        }
                      }}
                      value={fontFamily}
                    >
                      {fontChoices.map((font) => (
                        <option key={font} value={font}>
                          {font}
                        </option>
                      ))}
                    </select>
                  </label>

                  <label className="grid gap-2 text-sm font-medium text-slate-700">
                    Page
                    <select
                      className="rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-slate-950 outline-none transition focus:border-cyan-400 focus:bg-white focus:ring-4 focus:ring-cyan-100"
                      onChange={(event) => {
                        if (isPageSize(event.target.value)) {
                          setPageSize(event.target.value);
                        }
                      }}
                      value={pageSize}
                    >
                      {pageSizeChoices.map((size) => (
                        <option key={size.value} value={size.value}>
                          {size.label}
                        </option>
                      ))}
                    </select>
                  </label>
                </div>

                <label className="grid gap-3 text-sm font-medium text-slate-700">
                  Font size: {fontSize} pt
                  <input
                    className="accent-slate-950"
                    max={20}
                    min={9}
                    onChange={(event) => setFontSize(Number(event.target.value))}
                    type="range"
                    value={fontSize}
                  />
                </label>

                <label className="flex items-center justify-between gap-4 rounded-2xl bg-slate-50 px-4 py-3 text-sm font-medium text-slate-700">
                  Add centered title page heading
                  <input
                    checked={includeTitle}
                    className="h-5 w-5 accent-slate-950"
                    onChange={(event) => setIncludeTitle(event.target.checked)}
                    type="checkbox"
                  />
                </label>
              </div>
            </div>
          </aside>

          <section className="flex min-h-[40rem] flex-col rounded-[1.75rem] border border-slate-200 bg-slate-950 p-3 shadow-2xl shadow-slate-950/20">
            <div className="flex flex-wrap items-center justify-between gap-3 px-3 py-3 text-white">
              <div>
                <h2 className="text-lg font-semibold tracking-tight">Editor preview</h2>
                <p className="text-sm text-slate-400">{status}</p>
              </div>
              <button
                className="rounded-full bg-white px-5 py-3 text-sm font-semibold text-slate-950 shadow-lg shadow-white/10 transition hover:-translate-y-0.5 hover:bg-cyan-100 disabled:cursor-not-allowed disabled:opacity-50 disabled:hover:translate-y-0"
                disabled={!hasContent || isConverting}
                onClick={handleConvert}
                type="button"
              >
                {isConverting ? "Converting..." : "Download Word"}
              </button>
            </div>

            <div className="grid flex-1 overflow-hidden rounded-[1.25rem] bg-white lg:grid-cols-[1fr_18rem]">
              <textarea
                className="min-h-[32rem] resize-none bg-white p-6 font-mono text-sm leading-7 text-slate-800 outline-none lg:min-h-full"
                onChange={(event) => setText(event.target.value)}
                spellCheck={false}
                value={text}
              />

              <div className="border-t border-slate-200 bg-slate-50 p-5 lg:border-l lg:border-t-0">
                <div className="mb-5">
                  <p className="text-xs font-semibold uppercase tracking-[0.24em] text-slate-400">
                    Export name
                  </p>
                  <input
                    className="mt-2 w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold text-slate-950 outline-none transition focus:border-cyan-400 focus:ring-4 focus:ring-cyan-100"
                    onChange={(event) => setFileName(event.target.value)}
                    value={fileName}
                  />
                </div>

                <div className="space-y-3">
                  {[
                    "Real .docx output",
                    "Preserves line breaks",
                    "Works offline after load",
                    "No server upload",
                  ].map((feature) => (
                    <div
                      className="flex items-center gap-3 rounded-2xl border border-slate-200 bg-white p-3 text-sm font-medium text-slate-700"
                      key={feature}
                    >
                      <span className="grid h-6 w-6 place-items-center rounded-full bg-emerald-100 text-xs text-emerald-700">
                        ✓
                      </span>
                      {feature}
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </section>
        </div>
      </section>
    </main>
  );
}
