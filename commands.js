// Wordt door Office geladen vanuit function-file.html
Office.onReady(() => {
  // Niets speciaals nodig bij init.
});

// Deze functie is gekoppeld aan de lintknop via manifest.xml (onAction="insertHallo")
function insertHallo(event) {
  // Excel.run => context voor async Excel API
  Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.values = [["hallo"]]; // schrijf "hallo" in de actieve cel (of linksboven van selectie)
    await context.sync();
  })
  .catch((error) => {
    // Eventueel logging; in productie zou je dit netter afhandelen
    console.error("Fout bij insertHallo:", error);
  })
  .finally(() => {
    // VERY IMPORTANT: command-functies moeten het event.signaleren dat ze klaar zijn.
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  });
}

// Exporteer de functie-naam zodat Office deze kan vinden
if (typeof module !== "undefined") {
  module.exports = { insertHallo };
}
