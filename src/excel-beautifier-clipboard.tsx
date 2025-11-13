import { Detail, Clipboard, showToast, Toast, ActionPanel, Action } from "@raycast/api";
import { useEffect, useState } from "react";
import { ExcelFormulaBeautifier } from "./parser/excel-formula-parser";

export default function Command() {
  const [beautified, setBeautified] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    async function getClipboardContent() {
      try {
        const clipboardContent = await Clipboard.readText();
        if (!clipboardContent) {
          showToast({
            style: Toast.Style.Failure,
            title: "Error",
            message: "No text found in clipboard",
          });
          setIsLoading(false);
          return;
        }

        // Check if it looks like an Excel formula
        const trimmed = clipboardContent.trim();
        if (!trimmed.startsWith("=")) {
          showToast({
            style: Toast.Style.Animated,
            title: "Warning",
            message: "Content doesn't start with '='. Treating as formula anyway.",
          });
        }

        try {
          // Use our custom formatter with the new parser
          const result = ExcelFormulaBeautifier.rawText(trimmed);
          setBeautified(result);
          showToast({
            style: Toast.Style.Success,
            title: "Success",
            message: "Formula beautified",
          });
        } catch (formulaError) {
          showToast({
            style: Toast.Style.Failure,
            title: "Invalid Excel formula",
            message: formulaError instanceof Error ? formulaError.message : String(formulaError),
          });
        }
        setIsLoading(false);
      } catch (err) {
        showToast({
          style: Toast.Style.Failure,
          title: "Error",
          message: `Failed to read clipboard: ${err}`,
        });
        setIsLoading(false);
      }
    }

    getClipboardContent();
  }, []);

  if (isLoading) {
    return <Detail isLoading={true} markdown="Reading from clipboard..." />;
  }

  if (!beautified) {
    return <Detail markdown="# No formula data available" />;
  }

  const markdown = `## Beautified Formula
\`\`\`excel
${beautified}
\`\`\`
`;

  return (
    <Detail
      markdown={markdown}
      actions={
        <ActionPanel>
          <Action.CopyToClipboard
            title="Copy Beautified Formula"
            content={beautified}
            shortcut={{ modifiers: ["cmd"], key: "c" }}
          />
        </ActionPanel>
      }
    />
  );
}
