const reconcilerMap = {

    "GSC-BRZ": "Edson de Almeida",

    "GSC-SAN": "Josue Taveras",

    "GSC-EUC": "Kyle Rushton",

    "GSC-UKM": "Kyle Rushton",

    "GSC-ASN": "Myrna Guerrero",

    "GSC-PHL": "Myrna Guerrero",

    "GSC-ASI": "Nanette Davis",

    "GSC-MEX": "Nanette Davis",

    "GSC-AFC": "Robert Clark",

    "GSC-AFS": "Robert Clark",

    "GSC-AFW": "Robert Clark",

    "GSC-EUN": "Roberto Triumpho"
};

function showMessage(message) {

    document.getElementById("message")
        .innerText = message;
}

function validateRows(data) {

    let errors = [];

    data.forEach((row, index) => {

        const title =
            String(row.Title || "");

        const office =
            String(row.Office || "");

        const reconciler =
            String(row.Reconciler || "");

        const country =
            String(row.Country || "");

        if (!/^\d{4}/.test(title)) {

            errors.push(
                `Row ${index + 1}: Title must start with 4 numbers`
            );
        }

        if (office.trim() === "") {

            errors.push(
                `Row ${index + 1}: Office is empty`
            );
        }

        if (office === "GSC-CAM") {

            if (
                country.toUpperCase()
                    .includes("GUATEMALA")
            ) {

                if (
                    reconciler !==
                    "Mariuxi Maingon Andrade"
                ) {

                    errors.push(
                        `Row ${index + 1}: Guatemala in CAM must belong to Mariuxi`
                    );
                }

            } else {

                if (
                    reconciler !==
                    "Marcelo Correa"
                ) {

                    errors.push(
                        `Row ${index + 1}: CAM non Guatemala must belong to Marcelo`
                    );
                }
            }
        }

        if (office === "GSC-PAC") {

            if (
                country.toUpperCase()
                    .includes("AUS")
            ) {

                if (
                    reconciler !==
                    "Josue Taveras"
                ) {

                    errors.push(
                        `Row ${index + 1}: AUS in PAC must belong to Josue`
                    );
                }

            } else {

                if (
                    reconciler !==
                    "Edson de Almeida"
                ) {

                    errors.push(
                        `Row ${index + 1}: PAC non AUS must belong to Edson`
                    );
                }
            }
        }

        if (
            reconcilerMap[office] &&
            office !== "GSC-CAM" &&
            office !== "GSC-PAC"
        ) {

            if (
                reconciler !==
                reconcilerMap[office]
            ) {

                errors.push(
                    `Row ${index + 1}: ${office} must belong to ${reconcilerMap[office]}`
                );
            }
        }
    });

    return errors;
}

function processFile() {

    const fileInput =
        document.getElementById("sourceFile");

    const process =
        document.getElementById("processType")
            .value;

    const file =
        fileInput.files[0];

    if (!file) {

        showMessage(
            "Please select a file."
        );

        return;
    }

    const reader =
        new FileReader();

    reader.onload = function (e) {

        const data =
            new Uint8Array(e.target.result);

        const workbook =
            XLSX.read(data, {
                type: "array"
            });

        const sheet =
            workbook.Sheets[
                workbook.SheetNames[0]
            ];

        const json =
            XLSX.utils.sheet_to_json(sheet);

        const errors =
            validateRows(json);

        if (errors.length > 0) {

            showMessage(
                errors.slice(0, 20).join("\n")
            );

            return;
        }

        generateExcelFiles(
            json,
            process
        );

        showMessage(
            "Files generated successfully."
        );
    };

    reader.readAsArrayBuffer(file);
}

function generateExcelFiles(
    data,
    process
) {

    const workbook1 =
        XLSX.utils.book_new();

    const workbook2 =
        XLSX.utils.book_new();

    const sheet1 =
        XLSX.utils.json_to_sheet(data);

    const sheet2 =
        XLSX.utils.json_to_sheet(data);

    XLSX.utils.book_append_sheet(
        workbook1,
        sheet1,
        "Data"
    );

    XLSX.utils.book_append_sheet(
        workbook2,
        sheet2,
        "Data"
    );

    if (process === "NQ") {

        XLSX.writeFile(
            workbook1,
            "NQ by Accounts.xlsx"
        );

        XLSX.writeFile(
            workbook2,
            "Team Recon Report NQ.xlsx"
        );

    } else {

        XLSX.writeFile(
            workbook1,
            "Outstanding by Accounts.xlsx"
        );

        XLSX.writeFile(
            workbook2,
            "Team Recon Report Outstanding.xlsx"
        );
    }
}

function extractDateFromFilename(filename) {

    const monthMap = {

        jan: 0,
        january: 0,
        ene: 0,
        enero: 0,

        feb: 1,
        february: 1,
        febrero: 1,

        mar: 2,
        march: 2,
        marzo: 2,

        apr: 3,
        april: 3,
        abr: 3,
        abril: 3,

        may: 4,
        mayo: 4,

        jun: 5,
        june: 5,
        junio: 5,

        jul: 6,
        july: 6,
        julio: 6,

        aug: 7,
        august: 7,
        ago: 7,
        agosto: 7,

        sep: 8,
        september: 8,
        septiembre: 8,

        oct: 9,
        october: 9,
        octubre: 9,

        nov: 10,
        november: 10,
        noviembre: 10,

        dec: 11,
        december: 11,
        dic: 11,
        diciembre: 11
    };

    const regex =
        /(\d{1,2})\s([A-Za-z]+)\s(\d{4})/i;

    const match =
        filename.match(regex);

    if (!match) {

        throw new Error(
            `Date not recognized in ${filename}`
        );
    }

    const day =
        parseInt(match[1]);

    const monthText =
        match[2].toLowerCase();

    const year =
        parseInt(match[3]);

    const month =
        monthMap[monthText];

    if (month === undefined) {

        throw new Error(
            `Month not recognized in ${filename}`
        );
    }

    return new Date(
        year,
        month,
        day
    );
}

async function generateWeeklyReport() {

    const files =
        document.getElementById(
            "weeklyFiles"
        ).files;

    if (files.length === 0) {

        showMessage(
            "Please upload weekly files."
        );

        return;
    }

    let allData = [];

    for (const file of files) {

        const arrayBuffer =
            await file.arrayBuffer();

        const workbook =
            XLSX.read(arrayBuffer, {
                type: "array"
            });

        const sheet =
            workbook.Sheets[
                workbook.SheetNames[0]
            ];

        const json =
            XLSX.utils.sheet_to_json(sheet);

        const reportDate =
            extractDateFromFilename(
                file.name
            );

        allData.push({
            fileName: file.name,
            reportDate,
            data: json
        });
    }

    allData.sort(
        (a, b) =>
            a.reportDate - b.reportDate
    );

    let ranking = [];

    if (allData.length >= 2) {

        const first =
            allData[0].data;

        const last =
            allData[
                allData.length - 1
            ].data;

        const reconcilers =
            [
                ...new Set(
                    last.map(
                        x => x.Reconciler
                    )
                )
            ];

        reconcilers.forEach(
            reconciler => {

                const startItems =
                    first.filter(
                        x =>
                            x.Reconciler ===
                            reconciler
                    ).length;

                const endItems =
                    last.filter(
                        x =>
                            x.Reconciler ===
                            reconciler
                    ).length;

                const netReduction =
                    startItems -
                    endItems;

                ranking.push({

                    Reconciler:
                        reconciler,

                    "Week Start Items":
                        startItems,

                    "Week End Items":
                        endItems,

                    "Week Net Reduction":
                        netReduction
                });
            }
        );
    }

    ranking.sort(
        (a, b) =>
            b["Week Net Reduction"] -
            a["Week Net Reduction"]
    );

    ranking.forEach(
        (x, index) => {

            x.Rank = index + 1;
        }
    );

    const workbook =
        XLSX.utils.book_new();

    const rankingSheet =
        XLSX.utils.json_to_sheet(
            ranking
        );

    XLSX.utils.book_append_sheet(
        workbook,
        rankingSheet,
        "Weekly Ranking"
    );

    XLSX.writeFile(
        workbook,
        "Weekly Performance Report.xlsx"
    );

    showMessage(
        "Weekly report generated successfully."
    );
}