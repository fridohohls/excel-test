// Wartet, bis Office.js bereit ist
Office.onReady(function() {
    // Button-Event-Handler
    document.getElementById("colorButton").onclick = function() {
        Excel.run(function (context) {
            // Holt die aktuell markierte Zelle
            var range = context.workbook.getSelectedRange();
            range.format.fill.color = "red";
            return context.sync();
        }).catch(function(error) {
            console.log("Fehler: " + error);
        });
    };
});
