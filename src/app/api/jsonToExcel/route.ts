import { NextRequest, NextResponse } from "next/server";
import { google } from "googleapis";
import ExcelJS from "exceljs";
import fs from "fs";

// Load Google Service Account credentials
const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS || "{}"),
    scopes: ["https://www.googleapis.com/auth/drive"],
});
const drive = google.drive({ version: "v3", auth });

export async function POST(req: NextRequest) {
    try {
        const jsonData = await req.json();

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Data");

        // Define the column order
        const headers = [
            "id", "number", "date", "storage_id", "storage_name",
            "product_id", "product_name", "characteristic_name",
            "characteristic_id", "qty", "isNewProduct", "comment"
        ];
        worksheet.addRow(headers);

        // Function to add data
        const addProductData = (products: any[], isNewProduct = false) => {
            products.forEach((product) => {
                const { product: productInfo, characteristics, characteristic, qty } = product;
                const productId = productInfo.id;
                const productName = productInfo.name;

                if (characteristics) {
                    characteristics.forEach((char: any) => {
                        worksheet.addRow([
                            jsonData.id, jsonData.number, formatDate(jsonData.date),
                            jsonData.storage.id, jsonData.storage.name,
                            productId, productName, char.name, char.id,
                            char.qty, isNewProduct, jsonData.comment
                        ]);
                    });
                } else if (characteristic) {
                    worksheet.addRow([
                        jsonData.id, jsonData.number, formatDate(jsonData.date),
                        jsonData.storage.id, jsonData.storage.name,
                        productId, productName, characteristic.name,
                        characteristic.id, qty, isNewProduct, jsonData.comment
                    ]);
                }
            });
        };

        if (jsonData.products) addProductData(jsonData.products, false);
        if (jsonData.newproducts) addProductData(jsonData.newproducts, true);

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
