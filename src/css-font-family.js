// CSS identifiers containing digits (e.g. Source Sans 3) require quoting.
// Documents store family names, while CSS may contain a comma-separated stack.
export function cssFontFamily(value) {
  const generic = /^(?:serif|sans-serif|monospace|cursive|fantasy|system-ui|ui-serif|ui-sans-serif|ui-monospace|ui-rounded|emoji|math|fangsong)$/i;
  return String(value).split(',').map(part => part.trim()).filter(Boolean)
    .map(name => generic.test(name) ? name : JSON.stringify(name)).join(', ');
}
