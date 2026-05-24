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
  useState,
} from "react";

import {
  Accordion,
  AccordionContent,
  AccordionItem,
  AccordionTrigger,
} from "@/components/ui/accordion";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Separator } from "@/components/ui/separator";
import { Slider } from "@/components/ui/slider";
import { Switch } from "@/components/ui/switch";
import { Textarea } from "@/components/ui/textarea";

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

const seoHighlights = [
  {
    title: "Private TXT to Word conversion",
    description:
      "Your plain text file is read directly in the browser and converted locally, so the document content does not need to be uploaded to a server.",
  },
  {
    title: "Real DOCX output",
    description:
      "The tool creates a Word-compatible .docx file with title, font, page size, spacing, and preserved line breaks.",
  },
  {
    title: "Fast workflow for notes and drafts",
    description:
      "Convert notepad files, meeting notes, articles, reports, and clean text drafts into editable Word documents in seconds.",
  },
] as const;

const faqItems = [
  {
    question: "How do I convert TXT to Word?",
    answer:
      "Upload a .txt file in the hero section, review the live preview, choose your document settings, and click Download Word to save the DOCX file.",
  },
  {
    question: "Does the uploaded TXT file show a live preview?",
    answer:
      "Yes. After upload, the text appears instantly in the editor, the stats update, and the Word-style preview shows how the document will look.",
  },
  {
    question: "Are my files uploaded to a server?",
    answer:
      "No. File reading and DOCX generation happen in your browser tab, which keeps the conversion private.",
  },
  {
    question: "What file format does this tool download?",
    answer:
      "The converter downloads a modern .docx file that can be opened in Microsoft Word, Google Docs, LibreOffice, and other compatible editors.",
  },
] as const;

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
  const previewTitle = documentTitle.trim() || "Converted Document";
  const previewLines = useMemo(
    () => text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n").slice(0, 12),
    [text],
  );

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

  const handleFontSizeChange = (value: number | readonly number[]): void => {
    const nextFontSize = Array.isArray(value) ? value[0] : value;
    if (typeof nextFontSize === "number") {
      setFontSize(nextFontSize);
    }
  };

  return (
    <main className="min-h-screen overflow-hidden bg-[#0b1020] px-4 py-5 text-slate-950 sm:px-6 lg:px-8">
      <div className="pointer-events-none fixed inset-0 -z-10">
        <div className="absolute left-[-10%] top-[-15%] h-96 w-96 rounded-full bg-fuchsia-500/30 blur-3xl" />
        <div className="absolute right-[-8%] top-[12%] h-[28rem] w-[28rem] rounded-full bg-cyan-400/25 blur-3xl" />
        <div className="absolute bottom-[-18%] left-[30%] h-[30rem] w-[30rem] rounded-full bg-violet-500/20 blur-3xl" />
      </div>

      <section className="mx-auto w-full max-w-7xl rounded-[2rem] border border-white/15 bg-white/90 p-4 shadow-2xl shadow-black/30 backdrop-blur-2xl sm:p-6 lg:p-8">
        <header className="mb-8 flex flex-wrap items-center justify-between gap-4 rounded-[1.5rem] border border-slate-200 bg-white px-5 py-4 shadow-sm">
          <a className="flex items-center gap-3" href="#">
            <span className="grid h-11 w-11 place-items-center rounded-2xl bg-slate-950 text-sm font-black text-white">
              TW
            </span>
            <span>
              <span className="block text-base font-bold tracking-tight text-slate-950">
                TXT to Word
              </span>
              <span className="block text-xs font-medium text-slate-500">
                Browser DOCX converter
              </span>
            </span>
          </a>

          <nav className="flex items-center gap-2 text-sm font-semibold text-slate-600">
            <a className="rounded-full px-4 py-2 transition hover:bg-slate-100" href="#tool">
              Tool
            </a>
            <a className="rounded-full px-4 py-2 transition hover:bg-slate-100" href="#guide">
              Guide
            </a>
            <a className="rounded-full px-4 py-2 transition hover:bg-slate-100" href="#faq">
              FAQs
            </a>
          </nav>
        </header>

        <section
          className="grid gap-8 py-6 lg:grid-cols-[0.92fr_1.08fr] lg:items-center"
          id="tool"
        >
          <div>
            <Badge className="mb-5 gap-2 rounded-full border-slate-200 bg-white px-4 py-2 text-sm font-semibold text-slate-600 shadow-sm">
              <span className="h-2 w-2 rounded-full bg-emerald-400 shadow-[0_0_18px_rgba(52,211,153,0.9)]" />
              Free, private, no signup
            </Badge>
            <h1 className="max-w-4xl text-5xl font-semibold tracking-[-0.05em] text-slate-950 sm:text-6xl lg:text-7xl">
              Convert TXT to Word online.
            </h1>
            <p className="mt-5 max-w-2xl text-lg leading-8 text-slate-600">
              Upload a .txt file in this hero section, preview the content live,
              customize the document, and download a clean Word-compatible DOCX.
            </p>

            <div className="mt-7 grid grid-cols-3 gap-3 rounded-3xl border border-slate-200 bg-slate-950 p-3 text-white shadow-xl shadow-slate-950/20">
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
          </div>

          <Label
            htmlFor="txt-upload"
            className={`group flex min-h-96 cursor-pointer flex-col items-center justify-center rounded-[2rem] border border-dashed p-8 text-center transition ${
              isDragging
                ? "border-cyan-400 bg-cyan-50 shadow-2xl shadow-cyan-300/30"
                : "border-slate-300 bg-slate-50 hover:border-slate-400 hover:bg-white"
            }`}
            onDragEnter={() => setIsDragging(true)}
            onDragLeave={() => setIsDragging(false)}
            onDragOver={(event) => event.preventDefault()}
            onDrop={handleDrop}
          >
            <Input
              accept=".txt,text/plain"
              className="sr-only"
              id="txt-upload"
              onChange={handleFileChange}
              type="file"
            />
            <span className="mb-6 grid h-24 w-24 place-items-center rounded-[1.75rem] bg-slate-950 text-3xl font-black text-white shadow-2xl shadow-slate-950/30 transition group-hover:scale-105">
              TXT
            </span>
            <span className="text-3xl font-semibold tracking-tight text-slate-950">
              Upload TXT file
            </span>
            <span className="mt-3 max-w-md text-sm leading-6 text-slate-500">
              Drag and drop your .txt file here or click to browse. The live
              preview updates immediately after upload.
            </span>
            <span className="mt-6 rounded-full bg-primary px-6 py-3 text-sm font-semibold text-primary-foreground shadow-lg shadow-slate-950/20 transition group-hover:-translate-y-0.5">
              Select TXT File
            </span>
          </Label>
        </section>

        <section className="mt-8 grid gap-5 lg:grid-cols-[0.78fr_1.22fr]">
          <Card className="rounded-[1.75rem] border-slate-200 shadow-sm">
            <CardHeader className="flex flex-row items-center justify-between gap-4">
              <CardTitle className="text-lg tracking-tight text-slate-950">
                Document settings
              </CardTitle>
              <Badge className="rounded-full bg-emerald-50 px-3 py-1 text-xs font-semibold text-emerald-700">
                DOCX
              </Badge>
            </CardHeader>

            <CardContent className="grid gap-4">
              <div className="grid gap-2">
                <Label htmlFor="document-title">Document title</Label>
                <Input
                  className="h-11 rounded-2xl bg-slate-50 focus-visible:border-cyan-400 focus-visible:ring-cyan-100"
                  id="document-title"
                  onChange={(event) => setDocumentTitle(event.target.value)}
                  value={documentTitle}
                />
              </div>

              <div className="grid gap-4 sm:grid-cols-2">
                <div className="grid gap-2">
                  <Label>Font</Label>
                  <Select
                    onValueChange={(value) => {
                      if (value !== null && isFontChoice(value)) {
                        setFontFamily(value);
                      }
                    }}
                    value={fontFamily}
                  >
                    <SelectTrigger className="h-11 w-full rounded-2xl bg-slate-50">
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                    {fontChoices.map((font) => (
                      <SelectItem key={font} value={font}>
                        {font}
                      </SelectItem>
                    ))}
                    </SelectContent>
                  </Select>
                </div>

                <div className="grid gap-2">
                  <Label>Page</Label>
                  <Select
                    onValueChange={(value) => {
                      if (value !== null && isPageSize(value)) {
                        setPageSize(value);
                      }
                    }}
                    value={pageSize}
                  >
                    <SelectTrigger className="h-11 w-full rounded-2xl bg-slate-50">
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                    {pageSizeChoices.map((size) => (
                      <SelectItem key={size.value} value={size.value}>
                        {size.label}
                      </SelectItem>
                    ))}
                    </SelectContent>
                  </Select>
                </div>
              </div>

              <div className="grid gap-3">
                <Label>Font size: {fontSize} pt</Label>
                <Slider
                  max={20}
                  min={9}
                  onValueChange={handleFontSizeChange}
                  value={[fontSize]}
                />
              </div>

              <div className="flex items-center justify-between gap-4 rounded-2xl bg-slate-50 px-4 py-3">
                <Label htmlFor="include-title">Add centered title page heading</Label>
                <Switch
                  checked={includeTitle}
                  id="include-title"
                  onCheckedChange={setIncludeTitle}
                />
              </div>
            </CardContent>
          </Card>

          <section className="flex min-h-[40rem] flex-col rounded-[1.75rem] border border-slate-200 bg-slate-950 p-3 shadow-2xl shadow-slate-950/20">
            <div className="flex flex-wrap items-center justify-between gap-3 px-3 py-3 text-white">
              <div>
                <h2 className="text-lg font-semibold tracking-tight">Editor and live preview</h2>
                <p className="text-sm text-slate-400">{status}</p>
              </div>
              <Button
                className="rounded-full bg-white px-5 py-3 text-sm font-semibold text-slate-950 shadow-lg shadow-white/10 transition hover:-translate-y-0.5 hover:bg-cyan-100 disabled:hover:translate-y-0"
                disabled={!hasContent || isConverting}
                onClick={handleConvert}
                type="button"
              >
                {isConverting ? "Converting..." : "Download Word"}
              </Button>
            </div>

            <div className="grid flex-1 overflow-hidden rounded-[1.25rem] bg-white lg:grid-cols-[1fr_19rem]">
              <Textarea
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
                  <Input
                    className="mt-2 h-11 rounded-2xl bg-white text-sm font-semibold text-slate-950 focus-visible:border-cyan-400 focus-visible:ring-cyan-100"
                    onChange={(event) => setFileName(event.target.value)}
                    value={fileName}
                  />
                </div>

                <Card className="mb-5 rounded-3xl border-slate-200 shadow-sm">
                  <CardHeader className="flex flex-row items-center justify-between gap-3 p-4 pb-0">
                    <p className="text-xs font-semibold uppercase tracking-[0.24em] text-slate-400">
                      Live Word preview
                    </p>
                    <Badge className="rounded-full bg-cyan-50 px-2.5 py-1 text-[0.65rem] font-bold uppercase tracking-[0.16em] text-cyan-700">
                      Live
                    </Badge>
                  </CardHeader>

                  <CardContent className="p-4">
                  <div
                    className="max-h-80 overflow-hidden rounded-2xl border border-slate-100 bg-[#fbfbf8] px-5 py-6 text-slate-900 shadow-inner"
                    style={{ fontFamily, fontSize: `${fontSize}px` }}
                  >
                    {includeTitle ? (
                      <h3 className="mb-5 text-center text-lg font-bold leading-snug">
                        {previewTitle}
                      </h3>
                    ) : null}

                    <div className="space-y-2 leading-7">
                      {hasContent ? (
                        previewLines.map((line, index) => (
                          <p
                            className={line.trim().length === 0 ? "h-4" : "break-words"}
                            key={`${line}-${index}`}
                          >
                            {line}
                          </p>
                        ))
                      ) : (
                        <p className="text-slate-400">Upload or paste text to preview it here.</p>
                      )}
                    </div>
                  </div>
                  </CardContent>
                </Card>

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
                      <Badge className="grid h-6 w-6 place-items-center rounded-full bg-emerald-100 p-0 text-xs text-emerald-700">
                        ✓
                      </Badge>
                      {feature}
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </section>
        </section>

        <section className="mt-16 rounded-[2rem] border border-slate-200 bg-white p-6 shadow-sm lg:p-10" id="guide">
          <div className="max-w-3xl">
            <p className="text-sm font-bold uppercase tracking-[0.24em] text-cyan-700">
              TXT to Word guide
            </p>
            <h2 className="mt-3 text-4xl font-semibold tracking-[-0.04em] text-slate-950">
              Convert plain text into a professional Word document.
            </h2>
            <p className="mt-4 text-lg leading-8 text-slate-600">
              TXT files are simple and lightweight, but Word documents are better
              for sharing, editing, printing, and presenting. This converter
              turns raw text into a clean DOCX file while giving you control over
              the document title, font, page size, and preview.
            </p>
          </div>

          <div className="mt-8 grid gap-4 md:grid-cols-3">
            {seoHighlights.map((item) => (
              <Card className="rounded-3xl border-slate-200 bg-slate-50" key={item.title}>
                <CardHeader>
                  <CardTitle className="text-xl tracking-tight text-slate-950">
                    {item.title}
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <p className="leading-7 text-slate-600">{item.description}</p>
                </CardContent>
              </Card>
            ))}
          </div>

          <div className="mt-10 grid gap-8 lg:grid-cols-2">
            <div>
              <h3 className="text-2xl font-semibold tracking-tight text-slate-950">
                How to use this TXT to Word converter
              </h3>
              <ol className="mt-5 space-y-4 text-slate-600">
                {[
                  "Upload your .txt file in the hero upload box or paste text into the editor.",
                  "Check the live preview and adjust the document title, font, page size, and font size.",
                  "Click Download Word to generate and save your DOCX file instantly.",
                ].map((step, index) => (
                  <li className="flex gap-3 leading-7" key={step}>
                    <Badge className="grid h-8 w-8 shrink-0 place-items-center rounded-full bg-slate-950 p-0 text-sm font-bold text-white">
                      {index + 1}
                    </Badge>
                    <span>{step}</span>
                  </li>
                ))}
              </ol>
            </div>

            <div className="rounded-3xl bg-slate-950 p-6 text-white">
              <h3 className="text-2xl font-semibold tracking-tight">Best for</h3>
              <div className="mt-5 grid gap-3 sm:grid-cols-2">
                {[
                  "Notepad files",
                  "Meeting notes",
                  "Draft articles",
                  "Reports",
                  "Academic notes",
                  "Plain text exports",
                ].map((useCase) => (
                  <Badge className="justify-start rounded-2xl bg-white/10 p-4 text-sm font-medium text-white" key={useCase}>
                    {useCase}
                  </Badge>
                ))}
              </div>
            </div>
          </div>
        </section>

        <section className="mt-10 rounded-[2rem] border border-slate-200 bg-slate-50 p-6 lg:p-10" id="faq">
          <div className="max-w-3xl">
            <p className="text-sm font-bold uppercase tracking-[0.24em] text-cyan-700">
              Frequently asked questions
            </p>
            <h2 className="mt-3 text-4xl font-semibold tracking-[-0.04em] text-slate-950">
              TXT to Word FAQs
            </h2>
          </div>

          <Accordion className="mt-8 grid gap-4">
            {faqItems.map((item) => (
              <AccordionItem
                className="rounded-3xl border border-slate-200 bg-white px-6 shadow-sm"
                key={item.question}
                value={item.question}
              >
                <AccordionTrigger className="text-left text-lg font-semibold tracking-tight text-slate-950">
                  {item.question}
                </AccordionTrigger>
                <AccordionContent className="leading-7 text-slate-600">
                  {item.answer}
                </AccordionContent>
              </AccordionItem>
            ))}
          </Accordion>
        </section>

        <Separator className="mt-10" />
        <footer className="mt-6 flex flex-wrap items-center justify-between gap-4 text-sm text-slate-500">
          <p>© 2026 TXT to Word Converter. All rights reserved.</p>
          <div className="flex flex-wrap items-center gap-3">
            <a className="font-semibold text-slate-700 transition hover:text-slate-950" href="#tool">
              Convert TXT
            </a>
            <a className="font-semibold text-slate-700 transition hover:text-slate-950" href="#faq">
              FAQs
            </a>
            <a
              className="font-semibold text-slate-700 transition hover:text-slate-950"
              href="https://aslitools.com/tools/txt-to-word"
              rel="noreferrer"
              target="_blank"
            >
              Live tool
            </a>
          </div>
        </footer>
      </section>
    </main>
  );
}
