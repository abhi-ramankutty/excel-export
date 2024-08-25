const express = require("express");
const exceljs = require("exceljs");
const fs = require("fs");

const app = express();
const PORT = 4200;

app.get("/export", async (request, result) => {
    const workbook = new exceljs.Workbook();
    const sheet = workbook.addWorksheet("MutualFund");

    sheet.columns = [
        { header: "Name", key: "name" },
        { header: "SubSector", key: "subsector" },
        { header: "Option", key: "option" },
        { header: "AUM", key: "aum" },
        { header: "Expense Ratio", key: "expRatio" },
        { header: "Tracking Error", key: "trackErr" },
        { header: "5Y CAGR", key: "ret5y" },
    ];

    const dataObj = JSON.parse(fs.readFileSync("tickertape.json", "utf-8"));
    await dataObj.result.map((item) => {
        const values = item.values;

        const aum = values.find((a) => a.filter == "aum")?.doubleVal;
        const subsector = values.find((a) => a.filter == "subsector")?.strVal;
        const option = values.find((a) => a.filter == "option")?.strVal;
        const expRatio = values.find((a) => a.filter == "expRatio")?.doubleVal;
        const trackErr = values.find((a) => a.filter == "trackErr")?.doubleVal;
        const ret5y = values.find((a) => a.filter == "ret5y")?.doubleVal;

        sheet.addRow({
            name: item.name,
            aum,
            subsector,
            option,
            expRatio,
            ret5y,
            trackErr,
        });
    });

    result.setHeader(
        "Content-type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    result.setHeader(
        "Content-disposition",
        "attachment; filename=TickertapeMFScreener.xlsx"
    );

    workbook.xlsx.write(result);
});

app.listen(PORT, () => {
    console.log("App is running on http://localhost:4200");
    console.log("Navigate to http://localhost:4200/export to download your file");
});
