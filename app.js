const ExcelJS = require("exceljs");
const express = require("express");
const winston = require("winston");
const moment = require("moment-timezone");
const axios = require("axios");

var cors = require("cors");
const app = express();
const port = 3000;

// app.use(bodyParser.urlencoded({ extended: true }));
app.use(cors());
app.use(express.json());

const fetchAllowed = async () => {
    const response = await axios.get("https://repo-ctp.test/api/alloweds");
    return response.data;
};

app.post("/simpan-ke-xlsx", async (req, res) => {
    try {
        const ExcelJS = require("exceljs");
        const workbook = new ExcelJS.Workbook();
        const workbookBackup = new ExcelJS.Workbook();

        await workbook.xlsx.readFile("data.xlsx");
        await workbookBackup.xlsx.readFile("backup.xlsx");

        const worksheet =
            workbook.getWorksheet("Data Formulir") ||
            workbook.addWorksheet("Data Formulir");

        const worksheetBackup =
            workbookBackup.getWorksheet("Data Formulir") ||
            workbookBackup.addWorksheet("Data Formulir");

        const headers = [
            "AWB",
            "PACKAGE ID / NO ORDER",
            "NAMA PIC",
            "ACCOUNT",
            "ORIGIN",
            "REGIONAL",
            "NAMA KOTA",
            "CUSTOMER NAME",
            "SELLER NAME",
            "MERCHANT ID",
            "SELLER ADDRESS",
            "SELLER PHONE",
            "DESTINASI",
            "Service",
            "INTRUKSI",
            "Ins Value",
            "COD & NON COD",
            "COD AMOUNT",
            "ARMADA",
            "MODUL ENTRY",
            "KET",
            "QTY",
            "WEIGHT",
            "Create Time",
            "Start Time",
            "DISTRIK",
            "ZONA",
            "DATE REQUEST",
        ];

        if (worksheet.rowCount === 0) worksheet.addRow(headers);
        if (worksheetBackup.rowCount === 0) worksheetBackup.addRow(headers);

        // ================= HELPER =================

        const clean = (val = "") =>
            String(val)
                .toUpperCase()
                .trim()
                .replace(/\t/g, "")
                .replace(/\s+/g, " ");

        const splitOrSingle = (val) => (val ? String(val).split("|") : [""]);

        const today = new Date();
        const formattedDate = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;

        // ================= SPLIT ALL =================

        const fields = {};
        Object.keys(req.body).forEach((key) => {
            fields[key] = splitOrSingle(req.body[key]);
        });

        const totalRow = fields.code.length;

        // ================= LOOP =================

        let allowedAccounts = [];

        try {
            allowedAccounts = await fetchAllowed();
        } catch (err) {
            console.error("Gagal fetch allowed accounts:", err.message);
            return res.status(500).send("Gagal mengambil allowed accounts");
        }

        for (let i = 0; i < totalRow; i++) {
            const get = (key) => fields[key][i] ?? fields[key][0] ?? "";

            const no_orderNew = get("no_order");
            const seller_nameNew = get("seller_name");
            const accountNew = get("account");

            // // ================= ACCOUNT FILTER =================
            // const allowedAccounts = ["80640300", "80532006", "80090800"];

            // ================= PREPARE ORDER VALUE =================
            let orderValue = no_orderNew;

            if (orderValue?.includes("|")) {
                orderValue = orderValue.split("|")[0]?.trim();
            }

            if (orderValue) {
                orderValue = orderValue.replace(/\s+/g, "");
            }

            // ================= SELLER LOGIC =================
            let newSeller = seller_nameNew;

            if (
                allowedAccounts.includes(String(accountNew)) &&
                orderValue &&
                orderValue !== "-"
            ) {
                newSeller = `${orderValue}-${seller_nameNew}`;
            }

            const rowData = [
                clean(get("code")),
                clean(no_orderNew),
                clean(get("nama_pic")),
                clean(get("account")),
                clean(get("origin")).slice(0, 3),
                clean(get("regional")),
                clean(get("nama_kota")),
                clean(get("customer_name")),
                clean(newSeller),
                clean(get("merchant_id")),
                clean(get("seller_address")),
                clean(get("seller_phone")),
                clean(get("destinasi")),
                clean(get("service")),
                clean(get("intruksi")),
                clean(get("ins_value")),
                clean(get("cod_flag")),
                clean(get("cod_amount")),
                clean(get("armada")),
                clean(get("modul_entry")),
                clean(get("ket")),
                clean(get("qty")),
                clean(get("weight")),
                get("startime"),
                "",
                clean(get("nama_kota")),
                "",
                formattedDate,
            ];

            worksheet.addRow(rowData);
            worksheetBackup.addRow(rowData);
        }

        await workbook.xlsx.writeFile("data.xlsx");
        await workbookBackup.xlsx.writeFile("backup.xlsx");

        res.send("Data berhasil ditambahkan ke XLSX");
    } catch (error) {
        console.error("Error:", error.message);
        res.status(500).send("Terjadi kesalahan saat menambahkan data ke XLSX");
    }
});

app.post("/simpan-ke-xlsx-next", async (req, res) => {
    try {
        const ExcelJS = require("exceljs");
        const workbook = new ExcelJS.Workbook();
        const workbookBackup = new ExcelJS.Workbook();

        await workbook.xlsx.readFile("next.xlsx");
        await workbookBackup.xlsx.readFile("backup.xlsx");

        const worksheet =
            workbook.getWorksheet("Data Formulir") ||
            workbook.addWorksheet("Data Formulir");

        const worksheetBackup =
            workbookBackup.getWorksheet("Data Formulir") ||
            workbookBackup.addWorksheet("Data Formulir");

        const headers = [
            "AWB",
            "PACKAGE ID / NO ORDER",
            "NAMA PIC",
            "ACCOUNT",
            "ORIGIN",
            "REGIONAL",
            "NAMA KOTA",
            "CUSTOMER NAME",
            "SELLER NAME",
            "MERCHANT ID",
            "SELLER ADDRESS",
            "SELLER PHONE",
            "DESTINASI",
            "Service",
            "INTRUKSI",
            "Ins Value",
            "COD & NON COD",
            "COD AMOUNT",
            "ARMADA",
            "MODUL ENTRY",
            "KET",
            "QTY",
            "WEIGHT",
            "Create Time",
            "Start Time",
            "DISTRIK",
            "ZONA",
            "DATE REQUEST",
        ];

        if (worksheet.rowCount === 0) worksheet.addRow(headers);
        if (worksheetBackup.rowCount === 0) worksheetBackup.addRow(headers);

        // ================= HELPER =================

        const clean = (val = "") =>
            String(val)
                .toUpperCase()
                .trim()
                .replace(/\t/g, "")
                .replace(/\s+/g, " ");

        const splitOrSingle = (val) => (val ? String(val).split("|") : [""]);

        const today = new Date();
        const formattedDate = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;

        // ================= SPLIT ALL =================

        const fields = {};
        Object.keys(req.body).forEach((key) => {
            fields[key] = splitOrSingle(req.body[key]);
        });

        const totalRow = fields.code.length;

        // ================= LOOP =================

        let allowedAccounts = [];

        try {
            allowedAccounts = await fetchAllowed();
        } catch (err) {
            console.error("Gagal fetch allowed accounts:", err.message);
            return res.status(500).send("Gagal mengambil allowed accounts");
        }

        for (let i = 0; i < totalRow; i++) {
            const get = (key) => fields[key][i] ?? fields[key][0] ?? "";

            const no_orderNew = get("no_order");
            const seller_nameNew = get("seller_name");
            const accountNew = get("account");

            // // ================= ACCOUNT FILTER =================
            // const allowedAccounts = ["80640300", "80532006", "80090800"];

            // ================= PREPARE ORDER VALUE =================
            let orderValue = no_orderNew;

            if (orderValue?.includes("|")) {
                orderValue = orderValue.split("|")[0]?.trim();
            }

            if (orderValue) {
                orderValue = orderValue.replace(/\s+/g, "");
            }

            // ================= SELLER LOGIC =================
            let newSeller = seller_nameNew;

            if (
                allowedAccounts.includes(String(accountNew)) &&
                orderValue &&
                orderValue !== "-"
            ) {
                newSeller = `${orderValue}-${seller_nameNew}`;
            }

            const rowData = [
                clean(get("code")),
                clean(no_orderNew),
                clean(get("nama_pic")),
                clean(get("account")),
                clean(get("origin")).slice(0, 3),
                clean(get("regional")),
                clean(get("nama_kota")),
                clean(get("customer_name")),
                clean(newSeller),
                clean(get("merchant_id")),
                clean(get("seller_address")),
                clean(get("seller_phone")),
                clean(get("destinasi")),
                clean(get("service")),
                clean(get("intruksi")),
                clean(get("ins_value")),
                clean(get("cod_flag")),
                clean(get("cod_amount")),
                clean(get("armada")),
                clean(get("modul_entry")),
                clean(get("ket")),
                clean(get("qty")),
                clean(get("weight")),
                get("startime"),
                "",
                clean(get("nama_kota")),
                "",
                formattedDate,
            ];

            worksheet.addRow(rowData);
            worksheetBackup.addRow(rowData);
        }

        await workbook.xlsx.writeFile("data.xlsx");
        await workbookBackup.xlsx.writeFile("backup.xlsx");

        res.send("Data berhasil ditambahkan ke XLSX");
    } catch (error) {
        console.error("Error:", error.message);
        res.status(500).send("Terjadi kesalahan saat menambahkan data ke XLSX");
    }
});
app.post("/create-task-sca", async (req, res) => {
    try {
        console.log(req.body);

        const {
            code,
            no_order,
            nama_pic,
            nama_kota,
            regional,
            account,
            customer_name,
            origin,
            seller_name,
            merchant_id,
            seller_address,
            seller_phone,
            destinasi,
            service,
            intruksi,
            ins_value,
            armada,
            url,
        } = req.body;

        // ================= SPLIT DATA =================
        const split = (val, separator = "|") =>
            val ? val.split(separator) : [];

        const codeArray = split(code);
        const noOrderArray = split(no_order);
        const namaPicArray = split(nama_pic);
        const accountArray = split(account);
        const customerNameArray = split(customer_name);
        const originArray = split(origin);
        const sellerNameArray = split(seller_name);
        const merchantIdArray = merchant_id ? merchant_id.split("_") : [];
        const sellerAddressArray = split(seller_address);
        const sellerPhoneArray = seller_phone ? seller_phone.split("/") : [];
        const destinasiArray = split(destinasi);
        const serviceArray = split(service);
        const intruksiArray = split(intruksi);
        const insValueArray = split(ins_value);
        const armadaArray = split(armada);

        // ================= DATE & TIME =================
        const now = new Date();
        const formattedDate = now.toISOString().split("T")[0];
        const formattedTime = `${String(now.getHours()).padStart(2, "0")}:${String(
            now.getMinutes(),
        ).padStart(2, "0")}`;

        // ================= CLEANER FUNCTION =================
        const cleanText = (text = "") =>
            text
                .toUpperCase()
                .trim()
                .replace(/\t/g, "")
                .replace(/[^a-zA-Z0-9\s]/g, "")
                .replace(/\s+/g, " ");

        // ================= ACCOUNT FILTER =================
        let allowedAccounts = [];

        try {
            allowedAccounts = await fetchAllowed();
        } catch (err) {
            console.error("Gagal fetch allowed accounts:", err.message);
            return res.status(500).send("Gagal mengambil allowed accounts");
        }

        let orderValue = noOrderArray[0];

        if (orderValue?.includes("|")) {
            orderValue = orderValue.split("|")[0]?.trim();
        }

        if (orderValue) {
            orderValue = orderValue.replace(/\s+/g, "");
        }

        let pickupName = cleanText(sellerNameArray[0]).slice(0, 50);

        if (
            allowedAccounts.includes(String(accountArray[0])) &&
            orderValue &&
            orderValue !== "-"
        ) {
            pickupName = `${orderValue}-${cleanText(sellerNameArray[0])}`.slice(
                0,
                50,
            );
        }

        // ================= ORIGIN =================
        const originPrefix = cleanText(originArray[0]).slice(0, 3);
        const pickupBranch = `${originPrefix}000`;
        const pickupOrigin =
            originPrefix === "CGK" ? "CGK10000" : cleanText(originArray[0]);

        // ================= LOAD EXCEL =================
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile("upload_sca.xlsx");

        const sheet =
            workbook.getWorksheet("data") || workbook.addWorksheet("data");

        // ================= HEADER =================
        if (sheet.rowCount === 0) {
            sheet.addRow([
                "pickup_name",
                "pickup_customer",
                "pickup_courier_id",
                "pickup_status_latitude",
                "pickup_status_longitude",
                "pickup_branch",
                "pickup_origin",
                "pickup_zone",
                "pickup_service",
                "pickup_merchan_id",
                "pickup_address",
                "pickup_date",
                "pickup_time",
                "pickup_vehicle",
                "pickup_type",
                "pickup_pic",
                "pickup_pic_phone",
                "pickup_desc",
                "pickup_goods_desc",
                "pickup_doc_url",
                "pickup_cust_no",
            ]);
        }

        // ================= ADD ROW =================
        sheet.addRow([
            pickupName,
            cleanText(customerNameArray[0]),
            "",
            "",
            "",
            pickupBranch,
            pickupOrigin,
            "",
            cleanText(serviceArray[0]),
            merchantIdArray.slice(0, 2).join("_"),
            cleanText(sellerAddressArray[0]),
            formattedDate,
            formattedTime,
            cleanText(armadaArray[0]),
            "CORPORATE",
            cleanText(namaPicArray[0]),
            cleanText(sellerPhoneArray[0]).slice(0, 20),
            `${cleanText(intruksiArray[0])} + ${cleanText(insValueArray[0])}`,
            cleanText(destinasiArray[0]).slice(0, 60),
            url,
            cleanText(accountArray[0]),
        ]);

        await workbook.xlsx.writeFile("upload_sca.xlsx");

        res.send("Data berhasil ditambahkan ke XLSX");
    } catch (error) {
        console.error("Error:", error.message);
        res.status(500).send("Terjadi kesalahan saat menambahkan data ke XLSX");
    }
});

app.listen(port, () => {
    console.log(`Server berjalan di http://localhost:${port}`);
});
