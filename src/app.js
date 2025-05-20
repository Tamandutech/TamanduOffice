// Load our library that generates the document
const Docxtemplater = require("docxtemplater");
// Load PizZip library to load the docx/pptx/xlsx file in memory
const PizZip = require("pizzip");
// Load Excel reader
const XLSX = require('xlsx');
// Load LibreOffice converter
const libre = require('libreoffice-convert');

// Builtin file system utilities
const fs = require("fs");
const path = require("path");

const certificatePath = path.resolve(__dirname, "input/certificate.pptx");
const membersTablePath = path.resolve(__dirname, "input/members.xlsx");
const outputPath = path.resolve(__dirname, "output/certificate.pptx");

// Read the Excel file
const workbook = XLSX.readFile(membersTablePath);
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const members = XLSX.utils.sheet_to_json(worksheet);

console.log(members);

// Process a single member's certificate
async function generateCertificate(member) {
    // Load a fresh copy of the template for each member
    const content = fs.readFileSync(certificatePath, "binary");
    const zip = new PizZip(content);
    
    // Create a new document for each member
    const memberDoc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    /*
     * Render the presentation with member data
     */
    memberDoc.render({
        nome: member.NOME,
        RA: member.RA,
    });

    /*
     * Get the output presentation and export it as a Node.js buffer
     */
    const buf = memberDoc.toBuffer();

    try {
        // Convert to PDF
        const pdfBuf = await libre.convert(buf, '.pdf', undefined);
        
        // Write the PDF file
        const pdfOutputPath = path.resolve(__dirname, `output/certificate_${member.RA}.pdf`);
        fs.writeFileSync(pdfOutputPath, pdfBuf);
        
        console.log(`Generated certificate for ${member.NOME} (${member.RA})`);
        return { success: true, member };
    } catch (error) {
        console.error(`Error converting certificate for ${member.NOME} (${member.RA}):`, error);
        return { success: false, member, error };
    }
}

// Process items with a concurrency limit
async function processWithConcurrency(items, processFn, concurrencyLimit = 10) {
    const results = [];
    const chunks = [];
    
    // Split items into chunks of concurrencyLimit size
    for (let i = 0; i < items.length; i += concurrencyLimit) {
        chunks.push(items.slice(i, i + concurrencyLimit));
    }
    
    // Process each chunk
    for (let i = 0; i < chunks.length; i++) {
        console.log(`Processing batch ${i + 1}/${chunks.length} (${chunks[i].length} items)...`);
        const chunkResults = await Promise.all(
            chunks[i].map(item => processFn(item))
        );
        results.push(...chunkResults);
    }
    
    return results;
}

// Process all members with concurrency limit
async function generateAllCertificates() {
    console.log('Starting certificate generation with concurrency limit of 10...');
    const startTime = Date.now();
    
    const results = await processWithConcurrency(
        members,
        generateCertificate,
        10
    );
    
    const endTime = Date.now();
    const duration = (endTime - startTime) / 1000;
    
    // Log summary
    const successful = results.filter(r => r.success).length;
    const failed = results.filter(r => !r.success).length;
    
    console.log('\nCertificate Generation Summary:');
    console.log(`Total time: ${duration.toFixed(2)} seconds`);
    console.log(`Successfully generated: ${successful}`);
    console.log(`Failed: ${failed}`);
    
    if (failed > 0) {
        console.log('\nFailed certificates:');
        results
            .filter(r => !r.success)
            .forEach(r => console.log(`- ${r.member.NOME} (${r.member.RA}): ${r.error.message}`));
    }
}

// Run the certificate generation
generateAllCertificates().catch(error => {
    console.error('Error in certificate generation process:', error);
});

