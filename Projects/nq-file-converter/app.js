function processFile() {
    const fileInput = document.getElementById("fileInput");
    const processType = document.getElementById("processType").value;
    const status = document.getElementById("status");

    if (!fileInput.files.length) {
        alert("Please upload a file");
        return;
    }

    const file = fileInput.files[0];

    // ✅ Validate file name
    if (processType === "NQ" && !file.name.includes("- NQ") && !file.name.includes("- O")) {
        alert("Invalid file. Please upload a NQ file (-NQ or -O)");
        return;
    }

    if (processType === "OU" && !file.name.includes("- OU")) {
        alert("Invalid file. Please upload an Outstanding file (-OU)");
        return;
    }

    const reader = new FileReader();

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        if (!jsonData.length) {
            alert("Empty file");
            return;
        }

        // ✅ Transform data
        const processed = jsonData.map((row, index) => {
            return {
                "# Group ID": index + 1,
                "Title": row["Title"] || "",
                "Account": (row["Title"] || "").toString().substring(0, 4),
                "Amount": row["Amount"] || "",
                "Office": row["Office"] || "",
                "Status": row["Status"] || ""
            };
        });

        // =============================
        // OUTPUT FILE NAMES
        // =============================

        let file1Name = "";
        let file2Name = "";

        if (processType === "NQ") {
            file1Name = "Team Recon Report NQ.xlsx";
            file2Name = "NQ by Accounts.xlsx";
        } else {
            file1Name = "Team Recon Report Outstanding.xlsx";
            file2Name = "Outstanding by Accounts.xlsx";
        }

        // =============================
        // CREATE EXCEL FILES
        // =============================

        const ws1 = XLSX.utils.json_to_sheet(processed);
        const wb1 = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb1, ws1, "Report");

        XLSX.writeFile(wb1, file1Name);

        const ws2 = XLSX.utils.json_to_sheet(processed);
        const wb2 = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb2, ws2, "Accounts");

        XLSX.writeFile(wb2, file2Name);

        status.innerText = "Files generated successfully ✅";
    };

    reader.readAsArrayBuffer(file);
}