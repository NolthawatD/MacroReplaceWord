import mammoth from 'mammoth'; // A library to extract text from DOCX files
import path from "path";
import puppeteer from 'puppeteer';

const convertDocxToPDF = async (setFileName) => {
    const currentModuleUrl = new URL(import.meta.url);
    const currentModuleDir = path.dirname(currentModuleUrl.pathname);
    const twoStepsBack = path.join(currentModuleDir, "../..");
    const folderDownloadPath = path.join(twoStepsBack, "download");
    const filename = `${setFileName}.docx`;
    const localDocxPath = path.join(folderDownloadPath, "doc_demo1", filename);
    const tempPdfPath = path.join(folderDownloadPath, "pdf_demo1", `${filename.replace(".docx", ".pdf")}`);

    try {
        // Extract text content from the DOCX file
        const { value } = await mammoth.convertToHtml({ path: localDocxPath });
        const content = value; // You may need to process or clean up the content
        
        // Launch Puppeteer browser
		const browser = await puppeteer.launch({ args: ['--no-sandbox'] });
        const page = await browser.newPage();
        
        // Set content and generate PDF
        await page.setContent(content);
        await page.pdf({ path: tempPdfPath, format: 'A4' });
        
        // Close browser
        await browser.close();
        
        // Return the path to the generated PDF file
        return tempPdfPath;
    } catch (error) {
        console.error(error);
        throw new Error("Conversion failed");
    }
};

export { convertDocxToPDF };
