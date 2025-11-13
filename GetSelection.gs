function getSelection() {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) return 'No text selected.';
  
  const elements = selection.getRangeElements();
  let text = '';
  for (const el of elements) {
    if (el.getElement().editAsText) {
      const range = el.getElement().editAsText();
      text += range.getText().substring(el.getStartOffset(), el.getEndOffsetInclusive() + 1);
    }
  }
  return text.trim();
}
