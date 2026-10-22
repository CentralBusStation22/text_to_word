import { AlignmentType, Document, Packer, Paragraph, TextRun } from "docx";
import "./style.css";

const form = document.querySelector<HTMLFormElement>("#document-form");
const textInput = document.querySelector<HTMLTextAreaElement>("#text-input");
const stats = document.querySelector<HTMLElement>("#text-stats");
const clearButton = document.querySelector<HTMLButtonElement>("#clear-button");
const previewButton = document.querySelector<HTMLButtonElement>("#preview-button");
const previewPanel = document.querySelector<HTMLElement>("#document-preview");
const previewPage = document.querySelector<HTMLElement>("#preview-page");
const downloadButton = document.querySelector<HTMLButtonElement>("#download-button");
const downloadLabel = downloadButton?.querySelector<HTMLElement>(".download-label");
const formStatus = document.querySelector<HTMLElement>("#form-status");

if (
  !form ||
  !textInput ||
  !stats ||
  !clearButton ||
  !previewButton ||
  !previewPanel ||
  !previewPage ||
  !downloadButton ||
  !downloadLabel ||
  !formStatus
) {
  throw new Error("One or more document editor elements could not be found.");
}

const editorForm = form;
const editorTextInput = textInput;
const editorStats = stats;
const editorClearButton = clearButton;
const editorPreviewButton = previewButton;
const editorPreviewPanel = previewPanel;
const editorPreviewPage = previewPage;
const editorDownloadButton = downloadButton;
const editorDownloadLabel = downloadLabel;
const editorFormStatus = formStatus;

function updateEditorState(): void {
  const text = editorTextInput.value;
  const words = text.trim() ? text.trim().split(/\s+/u).length : 0;
  const characters = [...text].length;

  editorStats.textContent = `${words} ${words === 1 ? "מילה" : "מילים"} · ${characters} תווים`;
  editorClearButton.disabled = text.length === 0;
  editorPreviewButton.disabled = text.trim().length === 0;
  editorDownloadButton.disabled = text.trim().length === 0;
  editorFormStatus.textContent = "";

  if (!editorPreviewPanel.hidden) {
    renderDocumentPreview();
  }
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
    const lines = text.split(/\r\n|\r|\n/u);
    const headingIndex = lines.findIndex((line) =>
      /^\s{0,3}#{1,6}\s+\S/u.test(line),
    );

    if (headingIndex !== -1) {
      const heading = lines[headingIndex].match(
        /^\s{0,3}#{1,6}\s+(.+?)\s*#*\s*$/u,
      );

      if (heading) {
        return {
          title: heading[1].trim(),
          content: lines.slice(headingIndex + 1).join("\n").trim(),
          usedFallback: false,
        };
      }
    }

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
  return content.replace(/@\d+/gu, "").replace(/\s+([.,!?;:…])/gu, "$1");
}

function getCleanedDocumentParts(text: string): ExtractedDocument {
  const extracted = extractDocumentParts(text);
  return { ...extracted, content: cleanContent(extracted.content) };
}

function renderDocumentPreview(): void {
  const extracted = getCleanedDocumentParts(editorTextInput.value);
  editorPreviewPage.replaceChildren();

  if (extracted.title) {
    const title = window.document.createElement("h3");
    title.textContent = extracted.title;
    editorPreviewPage.append(title);
  }

  for (const line of extracted.content.split(/\r\n|\r|\n/u)) {
    const paragraph = window.document.createElement("p");
    paragraph.textContent = line || "\u00a0";
    editorPreviewPage.append(paragraph);
  }
}

editorTextInput.addEventListener("input", updateEditorState);

editorClearButton.addEventListener("click", () => {
  editorTextInput.value = "";
  editorTextInput.focus();
  updateEditorState();
});

editorPreviewButton.addEventListener("click", () => {
  const isOpening = editorPreviewPanel.hidden;
  editorPreviewPanel.hidden = !isOpening;
  editorPreviewButton.setAttribute("aria-expanded", String(isOpening));
  editorPreviewButton.querySelector("span")!.textContent = isOpening
    ? "סגירת תצוגה מקדימה"
    : "תצוגה מקדימה";

  if (isOpening) {
    renderDocumentPreview();
  }
});

editorForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  const text = editorTextInput.value;
  if (!text.trim()) {
    updateEditorState();
    editorTextInput.focus();
    return;
  }

  const extracted = getCleanedDocumentParts(text);
  if (!extracted.title && !extracted.content) {
    editorFormStatus.textContent = "לא נמצא טקסט אחרי הסימונים. כדאי לבדוק את הטקסט ולנסות שוב.";
    return;
  }

  editorDownloadButton.disabled = true;
  editorDownloadLabel.textContent = "מכינים את המסמך...";
  editorFormStatus.textContent = "";

  try {
    const paragraphs = [
      ...(extracted.title
        ? [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              bidirectional: true,
              spacing: { after: 280, line: 360 },
              children: [
                new TextRun({
                  text: extracted.title,
                  font: "Arial",
                  size: 36,
                  bold: true,
                  color: "283858",
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
              spacing: { after: 160, line: 336 },
              children: [
                new TextRun({
                  text: line || " ",
                  font: "Arial",
                  size: 24,
                  color: "333333",
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
    link.download = `${getSafeFileName(extracted.title)}.docx`;
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    editorFormStatus.textContent = extracted.usedFallback
      ? "לא נמצאו סימוני כותרת ותוכן; כל הטקסט נשמר כתוכן במסמך."
      : "המסמך מוכן וההורדה התחילה.";
  } catch (error) {
    console.error("Could not create the Word document.", error);
    editorFormStatus.textContent = "לא הצלחנו ליצור את המסמך. אפשר לנסות שוב.";
  } finally {
    editorDownloadButton.disabled = editorTextInput.value.trim().length === 0;
    editorDownloadLabel.textContent = "הורדת קובץ Word";
  }
});

updateEditorState();
