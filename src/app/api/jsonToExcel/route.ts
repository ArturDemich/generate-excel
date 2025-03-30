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
            "product_id", "product_name", "characteristic_name", 
            "characteristic_id", "qty", "isNewProduct",
            "id_doc", "number_doc", "date", "storage_id", "storage_name", "comment"
        ];

        worksheet.addRow(headers);

        // Function to add product data and color product_name
        const addProductData = (products: any[], isNewProduct = false) => {
            products.forEach((product) => {
                const { product: productInfo, characteristics, characteristic, qty } = product;
                const productId = productInfo.id;
                const productName = productInfo.name;

                if (characteristics) {
                    characteristics.forEach((char: any) => {
                        const row = worksheet.addRow([
                            productId, productName, char.name, char.id, 
                            char.qty, isNewProduct, "", "", "", 
                            jsonData.storage.id, jsonData.storage.name, jsonData.comment
                        ]);

                        // Set dark orange color for product_name if isNewProduct is true
                        if (isNewProduct) {
                            row.getCell(2).font = { color: { argb: 'FF8C00' } };  // Dark Orange for product_name (column 2)
                        }
                    });
                } else if (characteristic) {
                    const row = worksheet.addRow([
                        productId, productName, characteristic.name, characteristic.id, 
                        qty, isNewProduct, "", "", "", 
                        jsonData.storage.id, jsonData.storage.name, jsonData.comment
                    ]);

                    // Set dark orange color for product_name if isNewProduct is true
                    if (isNewProduct) {
                        row.getCell(2).font = { color: { argb: 'FF8C00' } };  // Dark Orange for product_name (column 2)
                    }
                }
            });
        };

        // Add products to sheet
        if (jsonData.products) addProductData(jsonData.products, false);
        if (jsonData.newproducts) addProductData(jsonData.newproducts, true);

        // Add static values for the first row
        const firstRow = worksheet.getRow(2);  // Second row, because the first is the header
        firstRow.getCell(7).value = jsonData.id;  // 'id_doc'
        firstRow.getCell(8).value = jsonData.number;  // 'number_doc'
        firstRow.getCell(9).value = formatDate(jsonData.date);  // 'date'
        firstRow.getCell(10).value = jsonData.storage.id;  // 'storage_id'
        firstRow.getCell(11).value = jsonData.storage.name;  // 'storage_name'
        firstRow.getCell(12).value = jsonData.comment;  // 'comment'

        const buffer = await workbook.xlsx.writeBuffer();
        const filePath = `/tmp/${jsonData.storage.name} ${formatDate(jsonData.date)}.xlsx`;

        fs.writeFileSync(filePath, Buffer.from(buffer));

        // Upload to Google Drive
        const driveResponse = await drive.files.create({
            requestBody: {
                name: `${jsonData.storage.name} ${formatDate(jsonData.date)}.xlsx`,
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
}

// Function to format date
function formatDate(dateStr: string) {
    const [year, day, month] = dateStr.split(/[-\s:]+/);
    return `${day}.${month}.${year}`;
}
