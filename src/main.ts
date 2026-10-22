import { AlignmentType, Document, Packer, Paragraph, TextRun } from "docx";
import "./style.css";

const form = document.querySelector<HTMLFormElement>("#document-form");
const textInput = document.querySelector<HTMLTextAreaElement>("#text-input");
const fileNameInput = document.querySelector<HTMLInputElement>("#file-name");
const stats = document.querySelector<HTMLElement>("#text-stats");
const clearButton = document.querySelector<HTMLButtonElement>("#clear-button");
const downloadButton = document.querySelector<HTMLButtonElement>("#download-button");
const downloadLabel = downloadButton?.querySelector<HTMLElement>(".download-label");
const formStatus = document.querySelector<HTMLElement>("#form-status");

if (
  !form ||
  !textInput ||
  !fileNameInput ||
  !stats ||
  !clearButton ||
  !downloadButton ||
  !downloadLabel ||
  !formStatus
) {
  throw new Error("One or more document editor elements could not be found.");
}

function updateEditorState(): void {
  const text = textInput.value;
  const words = text.trim() ? text.trim().split(/\s+/u).length : 0;
  const characters = [...text].length;

  stats.textContent = `${words} ${words === 1 ? "מילה" : "מילים"} · ${characters} תווים`;
  clearButton.disabled = text.length === 0;
  downloadButton.disabled = text.trim().length === 0;
  formStatus.textContent = "";
}

function getSafeFileName(value: string): string {
  const safeName = value
    .trim()
    .replace(/\.docx$/iu, "")
    .replace(/[<>:"/\\|?*\u0000-\u001f]/gu, "-")
    .replace(/[. ]+$/u, "")
    .trim();

  return safeName || "המסמך שלי";
}

interface ExtractedDocument {
  title: string;
  content: string;
  usedFallback: boolean;
}

function extractDocumentParts(text: string): ExtractedDocument {
  const titleMarker = "כותרת:";
  const contentMarker = "תוכן:";
  const titleMarkerIndex = text.indexOf(titleMarker);
  const contentMarkerIndex = text.indexOf(contentMarker);

  if (titleMarkerIndex === -1 && contentMarkerIndex === -1) {
    return { title: "", content: text.trim(), usedFallback: true };
  }

  if (titleMarkerIndex === -1) {
    return {
      title: "",
      content: text.slice(contentMarkerIndex + contentMarker.length).trim(),
      usedFallback: false,
    };
  }

  const titleStart = titleMarkerIndex + titleMarker.length;
  const hasContentAfterTitle = contentMarkerIndex > titleMarkerIndex;

  return {
    title: text
      .slice(titleStart, hasContentAfterTitle ? contentMarkerIndex : undefined)
      .trim(),
    content: hasContentAfterTitle
      ? text.slice(contentMarkerIndex + contentMarker.length).trim()
      : "",
    usedFallback: false,
  };
}

function cleanContent(content: string): string {
  return content.replace(/@\d+/gu, "");
}

textInput.addEventListener("input", updateEditorState);

clearButton.addEventListener("click", () => {
  textInput.value = "";
  textInput.focus();
  updateEditorState();
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const text = textInput.value;
  if (!text.trim()) {
    updateEditorState();
    textInput.focus();
    return;
  }

  const extracted = extractDocumentParts(text);
  extracted.content = cleanContent(extracted.content);
  if (!extracted.title && !extracted.content) {
    formStatus.textContent = "לא נמצא טקסט אחרי הסימונים. כדאי לבדוק את הטקסט ולנסות שוב.";
    return;
  }

  downloadButton.disabled = true;
  downloadLabel.textContent = "מכינים את המסמך...";
  formStatus.textContent = "";

  try {
    const paragraphs = [
      ...(extracted.title
        ? [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              bidirectional: true,
              spacing: { after: 200 },
              children: [
                new TextRun({
                  text: extracted.title,
                  font: "Arial",
                  size: 28,
                  bold: true,
                  rightToLeft: true,
                }),
              ],
            }),
          ]
        : []),
      ...extracted.content
        .split(/\r\n|\r|\n/u)
        .map(
          (line) =>
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              bidirectional: true,
              children: [
                new TextRun({
                  text: line || " ",
                  font: "Arial",
                  size: 24,
                  rightToLeft: true,
                }),
              ],
            }),
        ),
    ];

    const document = new Document({
      sections: [
        {
          properties: {
            page: {
              margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
            },
          },
          children: paragraphs,
        },
      ],
    });

    const blob = await Packer.toBlob(document);
    const url = URL.createObjectURL(blob);
    const link = window.document.createElement("a");
    link.href = url;
    link.download = `${getSafeFileName(fileNameInput.value)}.docx`;
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    formStatus.textContent = extracted.usedFallback
      ? "לא נמצאו סימוני כותרת ותוכן; כל הטקסט נשמר כתוכן במסמך."
      : "המסמך מוכן וההורדה התחילה.";
  } catch (error) {
    console.error("Could not create the Word document.", error);
    formStatus.textContent = "לא הצלחנו ליצור את המסמך. אפשר לנסות שוב.";
  } finally {
    downloadButton.disabled = textInput.value.trim().length === 0;
    downloadLabel.textContent = "הורדת קובץ Word";
  }
});

updateEditorState();
