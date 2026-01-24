const ExcelJS = require("exceljs");
const express = require("express");
const winston = require("winston");
const moment = require("moment-timezone");

var cors = require("cors");
const app = express();
const port = 3000;

// app.use(bodyParser.urlencoded({ extended: true }));
app.use(cors());
app.use(express.json());

app.post("/simpan-ke-xlsx", async (req, res) => {
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
    cod_flag,
    cod_amount,
    armada,
    modul_entry,
    ket,
    qty,
    weight,
    startime,
  } = req.body;

  const codeArray = code.split("|");
  const no_orderArray = no_order.split("|");
  const nama_picArray = nama_pic.split("|");
  const nama_kotaArray = nama_kota.split("|");
  const regionalArray = regional.split("|");
  const accountArray = account.split("|");
  const customer_nameArray = customer_name.split("|");
  const originArray = origin.split("|");
  const seller_nameArray = seller_name.split("|");
  const merchant_idArray = merchant_id.split("|");
  const seller_addressArray = seller_address.split("|");
  const seller_phoneArray = seller_phone.split("|");
  const destinasiArray = destinasi.split("|");
  const serviceArray = service.split("|");
  const intruksiArray = intruksi.split("|");
  const ins_valueArray = ins_value.split("|");
  const cod_flagArray = cod_flag.split("|");
  const cod_amountArray = cod_amount.split("|");
  const armadaArray = armada.split("|");
  const modul_entryArray = modul_entry.split("|");
  const ketArray = ket.split("|");
  const qtyArray = qty.split("|");
  const weightArray = weight.split("|");
  const startimeArray = startime.split("|");

  try {
    // Load the existing workbook
    const workbook = new ExcelJS.Workbook();
    const workbook1 = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("data.xlsx");
    await workbook1.xlsx.readFile("backup.xlsx");

    // Select or create a worksheet
    const worksheet =
      workbook.getWorksheet("Data Formulir") ||
      workbook.addWorksheet("Data Formulir");
    const worksheet1 =
      workbook1.getWorksheet("Data Formulir") ||
      workbook1.addWorksheet("Data Formulir");

    // Add header if the worksheet is newly created
    if (worksheet.rowCount === 0) {
      worksheet.addRow([
        "AWB",
        "PACKAGE ID / NO ORDER",
        "NAMA PIC ",
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
        "DATE 1# ATTEMPT",
      ]);
    }
    if (worksheet1.rowCount === 0) {
      worksheet1.addRow([
        "AWB",
        "PACKAGE ID / NO ORDER",
        "NAMA PIC ",
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
        "DATE 1# ATTEMPT",
      ]);
    }

    // Append data
    codeArray.map((code, index) => {
      let no_orderNew = no_orderArray[index] ? no_orderArray[index] : no_order;
      let nama_picNew = nama_picArray[index] ? nama_picArray[index] : nama_pic;
      let nama_kotaNew = nama_kotaArray[index]
        ? nama_kotaArray[index]
        : nama_kota;
      let regionalNew = regionalArray[index] ? regionalArray[index] : regional;
      let accountNew = accountArray[index] ? accountArray[index] : account;
      let customer_nameNew = customer_nameArray[index]
        ? customer_nameArray[0]
        : customer_name;
      let originNew = originArray[index] ? originArray[index] : origin;
      let seller_nameNew = seller_nameArray[index]
        ? seller_nameArray[index]
        : seller_name;
      let merchant_idNew = merchant_idArray[index]
        ? merchant_idArray[index]
        : merchant_id;
      let seller_addressNew = seller_addressArray[index]
        ? seller_addressArray[index]
        : seller_address;
      let seller_phoneNew = seller_phoneArray[index]
        ? seller_phoneArray[index]
        : seller_phone;
      let destinasiNew = destinasiArray[index]
        ? destinasiArray[index]
        : destinasi;
      let serviceNew = serviceArray[index] ? serviceArray[index] : service;
      let intruksiNew = intruksiArray[index] ? intruksiArray[index] : intruksi;
      let ins_valueNew = ins_valueArray[index]
        ? ins_valueArray[index]
        : ins_value;
      let cod_flagNew = cod_flagArray[index] ? cod_flagArray[index] : cod_flag;
      let cod_amountNew = cod_amountArray[index]
        ? cod_amountArray[index]
        : cod_amount;
      let armadaNew = armadaArray[index] ? armadaArray[index] : armada;
      let modul_entryNew = modul_entryArray[index]
        ? modul_entryArray[index]
        : modul_entry;
      let ketNew = ketArray[index] ? ketArray[index] : ket;
      let qtyNew = qtyArray[index] ? qtyArray[index] : qty;
      let weightNew = weightArray[index] ? weightArray[index] : weight;

      worksheet.addRow([
        code.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        no_orderNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        nama_picNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        accountNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        originNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " ")
          .slice(0, 3),
        regionalNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        nama_kotaNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        customer_nameNew.toUpperCase().trim(),
        seller_nameNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        merchant_idNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        seller_addressNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        seller_phoneNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        destinasiNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        serviceNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        intruksiNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        ins_valueNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        cod_flagNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        cod_amountNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        armadaNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        modul_entryNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        ketNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        qtyNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        weightNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        startime,
        "",
        nama_kotaNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        "",
        `${
          new Date().getMonth() + 1
        }/${new Date().getDate()}/${new Date().getFullYear()}`,
      ]);
      worksheet1.addRow([
        code.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        no_orderNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        nama_picNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        accountNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        originNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " ")
          .slice(0, 3),
        regionalNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        nama_kotaNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        customer_nameNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " ")
          .replace(/[^a-zA-Z0-9\s]/g, ""),
        seller_nameNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        merchant_idNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        seller_addressNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        seller_phoneNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        destinasiNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        serviceNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        intruksiNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        ins_valueNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        cod_flagNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        cod_amountNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        armadaNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        modul_entryNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        ketNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        qtyNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        weightNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        startime,
        "",
        nama_kotaNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        "",
        `${
          new Date().getMonth() + 1
        }/${new Date().getDate()}/${new Date().getFullYear()}`,
      ]);
    });

    // Save the updated workbook
    await workbook.xlsx.writeFile("data.xlsx");
    await workbook1.xlsx.writeFile("backup.xlsx");

    res.send("Data berhasil ditambahkan ke XLSX");
  } catch (error) {
    console.error("Error:", error.message);
    res.status(500).send("Terjadi kesalahan saat menambahkan data ke XLSX");
  }
});

app.post("/simpan-ke-xlsx-next", async (req, res) => {
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
    cod_flag,
    cod_amount,
    armada,
    modul_entry,
    ket,
    qty,
    weight,
    startime,
  } = req.body;

  const codeArray = code.split("|");
  const no_orderArray = no_order.split("|");
  const nama_picArray = nama_pic.split("|");
  const nama_kotaArray = nama_kota.split("|");
  const regionalArray = regional.split("|");
  const accountArray = account.split("|");
  const customer_nameArray = customer_name.split("|");
  const originArray = origin.split("|");
  const seller_nameArray = seller_name.split("|");
  const merchant_idArray = merchant_id.split("|");
  const seller_addressArray = seller_address.split("|");
  const seller_phoneArray = seller_phone.split("|");
  const destinasiArray = destinasi.split("|");
  const serviceArray = service.split("|");
  const intruksiArray = intruksi.split("|");
  const ins_valueArray = ins_value.split("|");
  const cod_flagArray = cod_flag.split("|");
  const cod_amountArray = cod_amount.split("|");
  const armadaArray = armada.split("|");
  const modul_entryArray = modul_entry.split("|");
  const ketArray = ket.split("|");
  const qtyArray = qty.split("|");
  const weightArray = weight.split("|");
  const startimeArray = startime.split("|");

  try {
    // Load the existing workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("next.xlsx");

    // Select or create a worksheet
    const worksheet =
      workbook.getWorksheet("Data Formulir") ||
      workbook.addWorksheet("Data Formulir");

    // Add header if the worksheet is newly created
    if (worksheet.rowCount === 0) {
      worksheet.addRow([
        "AWB",
        "PACKAGE ID / NO ORDER",
        "NAMA PIC ",
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
        "DATE 1# ATTEMPT",
      ]);
    }

    // Append data
    codeArray.map((code, index) => {
      let no_orderNew = no_orderArray[index] ? no_orderArray[index] : no_order;
      let nama_picNew = nama_picArray[index] ? nama_picArray[index] : nama_pic;
      let nama_kotaNew = nama_kotaArray[index]
        ? nama_kotaArray[index]
        : nama_kota;
      let regionalNew = regionalArray[index] ? regionalArray[index] : regional;
      let accountNew = accountArray[index] ? accountArray[index] : account;
      let customer_nameNew = customer_nameArray[index]
        ? customer_nameArray[0]
        : customer_name;
      let originNew = originArray[index] ? originArray[index] : origin;
      let seller_nameNew = seller_nameArray[index]
        ? seller_nameArray[index]
        : seller_name;
      let merchant_idNew = merchant_idArray[index]
        ? merchant_idArray[index]
        : merchant_id;
      let seller_addressNew = seller_addressArray[index]
        ? seller_addressArray[index]
        : seller_address;
      let seller_phoneNew = seller_phoneArray[index]
        ? seller_phoneArray[index]
        : seller_phone;
      let destinasiNew = destinasiArray[index]
        ? destinasiArray[index]
        : destinasi;
      let serviceNew = serviceArray[index] ? serviceArray[index] : service;
      let intruksiNew = intruksiArray[index] ? intruksiArray[index] : intruksi;
      let ins_valueNew = ins_valueArray[index]
        ? ins_valueArray[index]
        : ins_value;
      let cod_flagNew = cod_flagArray[index] ? cod_flagArray[index] : cod_flag;
      let cod_amountNew = cod_amountArray[index]
        ? cod_amountArray[index]
        : cod_amount;
      let armadaNew = armadaArray[index] ? armadaArray[index] : armada;
      let modul_entryNew = modul_entryArray[index]
        ? modul_entryArray[index]
        : modul_entry;
      let ketNew = ketArray[index] ? ketArray[index] : ket;
      let qtyNew = qtyArray[index] ? qtyArray[index] : qty;
      let weightNew = weightArray[index] ? weightArray[index] : weight;

      worksheet.addRow([
        code.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        no_orderNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        nama_picNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        accountNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        originNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " ")
          .slice(0, 3),
        regionalNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        nama_kotaNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        customer_nameNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " ")
          .replace(/[^a-zA-Z0-9\s]/g, ""),
        seller_nameNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        merchant_idNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        seller_addressNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        seller_phoneNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        destinasiNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        serviceNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        intruksiNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        ins_valueNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        cod_flagNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        cod_amountNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        armadaNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        modul_entryNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        ketNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        qtyNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        weightNew.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
        startime,
        "",
        nama_kotaNew
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
        "",
        `${
          new Date().getMonth() + 1
        }/${new Date().getDate()}/${new Date().getFullYear()}`,
      ]);
    });

    const logger = winston.createLogger({
      format: winston.format.combine(
        winston.format.timestamp({
          format: () => {
            return moment().tz("Asia/Jakarta").format("YYYY-MM-DD HH:mm:ss");
          },
        }),
        winston.format.ms(),
        winston.format.json()
      ),
      transports: [
        new winston.transports.File({
          filename: "application.log",
        }),
      ],
    });

    logger.log({
      level: "info",
      message: {
        code,
        account,
        origin,
        regional,
        nama_kota,
        customer_name,
        seller_name,
        merchant_id,
        seller_address,
        seller_phone,
        destinasi,
        service,
        intruksi,
        ins_value,
        cod_flag,
        cod_amount,
        armada,
        modul_entry,
        ket,
        qty,
        weight,
      },
    });

    // Save the updated workbook
    await workbook.xlsx.writeFile("next.xlsx");

    res.send("Data berhasil ditambahkan ke XLSX");
  } catch (error) {
    console.error("Error:", error.message);
    res.status(500).send("Terjadi kesalahan saat menambahkan data ke XLSX");
  }
});

app.post("/create-task", async (req, res) => {
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
    cod_flag,
    cod_amount,
    armada,
    modul_entry,
    ket,
    qty,
    weight,
    startime,
  } = req.body;

  const codeArray = code.split("|");
  const no_orderArray = no_order.split("|");
  const nama_picArray = nama_pic.split("|");
  const nama_kotaArray = nama_kota.split("|");
  const regionalArray = regional.split("|");
  const accountArray = account.split("|");
  const customer_nameArray = customer_name.split("|");
  const originArray = origin.split("|");
  const seller_nameArray = seller_name.split("|");
  const merchant_idArray = merchant_id.split("|");
  const seller_addressArray = seller_address.split("|");
  const seller_phoneArray = seller_phone.split("|");
  const destinasiArray = destinasi.split("|");
  const serviceArray = service.split("|");
  const intruksiArray = intruksi.split("|");
  const ins_valueArray = ins_value.split("|");
  const cod_flagArray = cod_flag.split("|");
  const cod_amountArray = cod_amount.split("|");
  const armadaArray = armada.split("|");
  const modul_entryArray = modul_entry.split("|");
  const ketArray = ket.split("|");
  const qtyArray = qty.split("|");
  const weightArray = weight.split("|");
  const startimeArray = startime.split("|");

  try {
    // Load the existing workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("upload.xlsx");

    // Select or create a worksheet
    const MainSheet =
      workbook.getWorksheet("Main") || workbook.addWorksheet("Main");
    const TaskLabelSheet =
      workbook.getWorksheet("Task Label") ||
      workbook.addWorksheet("Task Label");
    const PackageListSheet =
      workbook.getWorksheet("Package List") ||
      workbook.addWorksheet("Package List");

    function generateRandomString(length) {
      const characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
      let randomString = "";
      for (let i = 0; i < length; i++) {
        randomString += characters.charAt(
          Math.floor(Math.random() * characters.length)
        );
      }
      return randomString;
    }

    // Add header if the worksheet is newly created
    if (MainSheet.rowCount === 0) {
      MainSheet.addRow([
        "Task Ref Id",
        "ID Rider",
        "Task Caption",
        "Customer Name",
        "Address",
        "Extra Detail",
        "PUORDER_PHONE",
        "PUORDER_BRANCH",
        "PUORDER_CUST_ID",
        "PUORDER_WEIGHT",
        "PUORDER_SPEC_INS",
        "PUORDER_PAYMENT_NAME",
        "PUORDER_SERVICES_NAME",
        "PUORDER_CUST_NAME",
        "PUORDER_TRANSMODE_NAME",
        "PUORDER_CITY",
        "PUORDER_GOODS_NAME",
        "PUORDER_QTY",
        "Nama Kota",
        "PUORDER MERCHANT ID",
        "Email",
        "PUORDER_CONTACT",
      ]);
    }

    if (TaskLabelSheet.rowCount === 0) {
      TaskLabelSheet.addRow(["Label Name", "Label Colour", "Task Ref Id"]);
    }
    if (PackageListSheet.rowCount === 0) {
      PackageListSheet.addRow(["Airwaybill Number", "Task Ref Id"]);
    }

    function generateUniqueCode() {
      // Format CT08
      let code = "CT08PRODUCT";

      // Random String
      const randomString = generateRandomString(4);
      code += randomString;

      // Tanggal + Bulan + Tahun
      const currentDate = new Date();
      const monthAbbreviations = [
        "JAN",
        "FEB",
        "MAR",
        "APR",
        "MEI",
        "JUN",
        "JUL",
        "AGT",
        "SEP",
        "OKT",
        "NOV",
        "DES",
      ];
      const datePart =
        currentDate.getDate().toString().padStart(2, "0") +
        monthAbbreviations[currentDate.getMonth()] +
        currentDate.getFullYear();
      code += datePart;

      // Nomor Urut (bisa di-generate sesuai kebutuhan)
      // Misalnya, kita tambahkan nomor urut secara sederhana
      // Jika butuh lebih kompleks, sesuaikan dengan kebutuhan sistem Anda
      const noUrut = Math.floor(Math.random() * 10000)
        .toString()
        .padStart(4, "0");
      code += noUrut;

      return code;
    }

    // Contoh penggunaan
    const refID = generateUniqueCode();

    MainSheet.addRow([
      refID,
      "",
      customer_nameArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/\s+/g, " "),
      seller_nameArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/\s+/g, " "),
      seller_addressArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/\s+/g, " "),
      "SENT TO " +
        destinasiArray[0]
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
      seller_phoneArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/\s+/g, " "),
      originArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/\s+/g, " ")
        .slice(0, 3),
      "NA " +
        accountArray[0]
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
      weightArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/\s+/g, " "),
      intruksiArray[0] +
        " + " +
        ins_valueArray[0]
          .toUpperCase()
          .trim()
          .replace(/\t/g, "")
          .replace(/\s+/g, " "),
      cod_flagArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/\s+/g, " "),
      serviceArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/\s+/g, " "),
      no_orderArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/\s+/g, " "),
      armadaArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/\s+/g, " "),
      nama_kotaArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/\s+/g, " "),
      ketArray[0].toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
      qtyArray[0].toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
      nama_kotaArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/\s+/g, " "),
      "ccc.ct08@jne.co.id",
      "ccc.ct08@jne.co.id",
      nama_picArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/\s+/g, " "),
    ]);

    TaskLabelSheet.addRow([
      armada.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
      "#02a8f3",
      refID,
    ]);

    TaskLabelSheet.addRow([
      service.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " "),
      "#02a8f3",
      refID,
    ]);

    // // Append data
    codeArray.map((code, index) => {
      PackageListSheet.addRow([
        code
          ? code.toUpperCase().trim().replace(/\t/g, "").replace(/\s+/g, " ")
          : "-",
        refID,
      ]);
    });

    // Save the updated workbook
    await workbook.xlsx.writeFile("upload.xlsx");

    res.send("Data berhasil ditambahkan ke XLSX");
  } catch (error) {
    console.error("Error:", error.message);
    res.status(500).send("Terjadi kesalahan saat menambahkan data ke XLSX");
  }
});

app.post("/samsung-parse", async (req, res) => {
  const { q } = req.body;

  const data = q.split("\n");

  console.log(data);

  let Messages = data.map((element) => element.toUpperCase());

  try {
    // Load the existing workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("samsung.xlsx");

    // Select or create a worksheet
    const worksheet =
      workbook.getWorksheet("Data Samsung") ||
      workbook.addWorksheet("Data Samsung");

    // Save to Excel
    worksheet.addRow(Messages);

    // Save the updated workbook
    await workbook.xlsx.writeFile("samsung.xlsx");

    res.send("Data berhasil ditambahkan ke XLSX");
  } catch (error) {
    console.error("Error:", error.message);
    res.status(500).send("Terjadi kesalahan saat menambahkan data ke XLSX");
  }
});

app.post("/create-task-sca", async (req, res) => {
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
    cod_flag,
    cod_amount,
    armada,
    modul_entry,
    ket,
    qty,
    weight,
    startime,
    url,
  } = req.body;

  const codeArray = code.split("|");
  const no_orderArray = no_order.split("|");
  const nama_picArray = nama_pic.split("|");
  const nama_kotaArray = nama_kota.split("|");
  const regionalArray = regional.split("|");
  const accountArray = account.split("|");
  const customer_nameArray = customer_name.split("|");
  const originArray = origin.split("|");
  const seller_nameArray = seller_name.split("|");
  const merchant_idArray = merchant_id.split("|");
  const seller_addressArray = seller_address.split("|");
  const seller_phoneArray = seller_phone.split("/");
  const destinasiArray = destinasi.split("|");
  const serviceArray = service.split("|");
  const intruksiArray = intruksi.split("|");
  const ins_valueArray = ins_value.split("|");
  const cod_flagArray = cod_flag.split("|");
  const cod_amountArray = cod_amount.split("|");
  const armadaArray = armada.split("|");
  const modul_entryArray = modul_entry.split("|");
  const ketArray = ket.split("|");
  const qtyArray = qty.split("|");
  const weightArray = weight.split("|");
  const startimeArray = startime.split("|");

  const today = new Date();
  const formattedDate = today.toISOString().split("T")[0];
  const hours = String(today.getHours()).padStart(2, "0");
  const minutes = String(today.getMinutes()).padStart(2, "0");
  const formattedTime = `${hours}:${minutes}`;

  function generateRandomString(length) {
    const characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
    let randomString = "";
    for (let i = 0; i < length; i++) {
      randomString += characters.charAt(
        Math.floor(Math.random() * characters.length)
      );
    }
    return randomString;
  }

  function generateUniqueCode() {
    // Format CT08
    let code = "CT08PRO";
    code += customer_nameArray[0];

    // Random String
    const randomString = generateRandomString(4);

    // Tanggal + Bulan + Tahun
    const currentDate = new Date();
    const monthAbbreviations = [
      "JAN",
      "FEB",
      "MAR",
      "APR",
      "MEI",
      "JUN",
      "JUL",
      "AGT",
      "SEP",
      "OKT",
      "NOV",
      "DES",
    ];
    const datePart =
      currentDate.getDate().toString().padStart(2, "0") +
      monthAbbreviations[currentDate.getMonth()] +
      currentDate.getFullYear();
    code += datePart;

    code += randomString;

    // Nomor Urut (bisa di-generate sesuai kebutuhan)
    // Misalnya, kita tambahkan nomor urut secara sederhana
    // Jika butuh lebih kompleks, sesuaikan dengan kebutuhan sistem Anda
    const noUrut = Math.floor(Math.random() * 10000)
      .toString()
      .padStart(4, "0");
    code += noUrut;

    return code;
  }

  // Contoh penggunaan
  const refID = generateUniqueCode();

  try {
    // Load the existing workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("upload_sca.xlsx");

    // Select or create a worksheet
    const MainSheet =
      workbook.getWorksheet("data") || workbook.addWorksheet("data");

    // Add header if the worksheet is newly created
    if (MainSheet.rowCount === 0) {
      MainSheet.addRow([
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

    MainSheet.addRow([
      seller_nameArray[0]
        .toUpperCase()
        .trim()
        .replace(/\s+/g, " ")
        .slice(0, 50),
      customer_nameArray[0].toUpperCase().trim().replace(/\s+/g, " "),
      "",
      "",
      "",
      originArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/[^a-zA-Z0-9\s]/g, "")
        .replace(/\s+/g, " ")
        .slice(0, 3) + "000",
      originArray[0].slice(0, 3) == "CGK"
        ? "CGK10000"
        : originArray[0].toUpperCase().trim(),
      "",
      serviceArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/[^a-zA-Z0-9\s]/g, "")
        .replace(/\s+/g, " "),
      accountArray[0] + "_" + refID.trim().replace(/\s+/g, ""),
      seller_addressArray[0].toUpperCase().trim().replace(/\t/g, ""),
      formattedDate,
      formattedTime,
      armadaArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/[^a-zA-Z0-9\s]/g, "")
        .replace(/\s+/g, " "),
      "CORPORATE",
      nama_picArray[0].toUpperCase().trim().replace(/\s+/g, " "),
      seller_phoneArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/[^a-zA-Z0-9\s]/g, "")
        .replace(/\s+/g, " ")
        .slice(0, 20),
      intruksiArray[0].toUpperCase().trim().replace(/\t/g, "") +
        " + " +
        ins_valueArray[0].toUpperCase().trim().replace(/\t/g, ""),
      destinasiArray[0].toUpperCase().trim().slice(0, 60),
      //   ketArray[0].toUpperCase().trim().replace(/\t/g, "").slice(0, 60),
      //   "SENT TO " + destinasiArray[0].toUpperCase().trim().replace(/\t/g, ""),
      url,
      accountArray[0]
        .toUpperCase()
        .trim()
        .replace(/\t/g, "")
        .replace(/[^a-zA-Z0-9\s]/g, "")
        .replace(/\s+/g, " "),
    ]);

    // Save the updated workbook
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
