import { NextRequest, NextResponse } from "next/server";
import { google } from "googleapis";
import ExcelJS from "exceljs";
import fs from "fs";

// Load Google Service Account credentials
const credentialsBase64 = process.env.GOOGLE_CREDENTIALS;

const credentials = credentialsBase64 && JSON.parse(Buffer.from(credentialsBase64, 'base64').toString('utf-8'));
const auth = new google.auth.GoogleAuth({
    credentials: credentials,
    scopes: ["https://www.googleapis.com/auth/drive"],
});
const drive = google.drive({ version: "v3", auth });

export async function POST(req: NextRequest) {
    try {
        const jsonData = await req.json();

        // Validate required fields
        if (!jsonData || !jsonData.products || !jsonData.newproducts || !jsonData.storage || !jsonData.date) {
            return new NextResponse("Missing required data", { status: 400 });
        }

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Data");

        // Correct headers as per your requirement
        const headers = [
            "product_name",
            "characteristic_name",
            "qty",
            "saleQty",
            "unit_name",
            "plantComment",
            "isNewProduct",
            "storage_name",
            "product_id",
            "characteristic_id",
            "unit_id",
            "doc_comment",
            "date",
            "number_doc", "storage_id", "id_doc"
        ];

        worksheet.addRow(headers);

        // Function to add product data and color product_name
        const addProductData = (products: any[], isNewProduct = false, storageName: string) => {
            products.forEach((product) => {
                const { product: productInfo, characteristic, unit, qty, saleQty, plantComment } = product;
                const productId = productInfo.id;
                const productName = productInfo.name;

                if (characteristic) {
                    const row = worksheet.addRow([
                        productName || "",
                        characteristic.name || "",
                        qty,
                        saleQty,
                        unit?.name || "", 
                        plantComment || "",
                        isNewProduct,
                        storageName,
                        productId || "",
                        characteristic.id || "",
                        unit?.id || "", 
                        "", "", "", "", "", "",
                    ]);
                    row.getCell(4).font = { color: { argb: 'FF00A100' } };

                    // Set dark orange color for product_name if isNewProduct is true
                    if (isNewProduct) {
                        row.getCell(1).font = { color: { argb: 'FF8C00' } };
                        row.getCell(2).font = { color: { argb: 'FF8C00' } };
                    }
                }
            });
        };

        // Add products to sheet
        if (jsonData.products) addProductData(jsonData.products, false, jsonData.storage.name);
        if (jsonData.newproducts) addProductData(jsonData.newproducts, true, jsonData.storage.name);

        // Add static values for the first row
        const firstRow = worksheet.getRow(2);  // Second row, because the first is the header
        firstRow.getCell(12).value = jsonData.comment;  // 'comment'
        firstRow.getCell(13).value = formatDateTime(jsonData.date);  // 'date'
        firstRow.getCell(14).value = jsonData.number;  // 'number_doc'
        firstRow.getCell(15).value = jsonData.storage.id;  // 'storage_id'
        firstRow.getCell(16).value = jsonData.id;  // 'id_doc'

        const columnsToFit = [9, 10, 11, 15, 16];
        columnsToFit.forEach((colIndex) => {
            worksheet.getColumn(colIndex).hidden = true;
        });

        const buffer = await workbook.xlsx.writeBuffer();
        const filePath = `/tmp/${jsonData.storage.name} ${formatDateTime(jsonData.date)}.xlsx`;

        fs.writeFileSync(filePath, Buffer.from(buffer));

        // Upload to Google Drive
        const driveResponse = await drive.files.create({
            requestBody: {
                name: `${jsonData.storage.name} ${formatDateTime(jsonData.date)}.xlsx`,
                parents: [process.env.GOOGLE_DRIVE_FOLDER_ID || ""], // Google Drive folder ID
            },
            media: {
                mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                body: fs.createReadStream(filePath),
            },
            fields: "id",
        });

        fs.unlinkSync(filePath); // Cleanup after upload

        return NextResponse.json({
            message: "File uploaded to Google Drive",
            fileId: driveResponse.data.id,
        });

    } catch (error: any) {
        console.error("Error:", error);
        return new NextResponse(`Error: ${error.message}`, { status: 500 });
    }
};

// Function to format date

function formatDateTime(dateStr: string) {
    const [year, day, month] = dateStr.split(/[-\s:]+/);
    const time = new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
    return `${day}.${month}.${year}-${time}`;
};
