// server.js - FINAL VERSION (Sends Correctly Batched Emails to Faculty)

// 1. Import necessary packages
const express = require('express');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const cors = require('cors');
const nodemailer = require('nodemailer');
const cron = require('node-cron');

require('dotenv').config();

// 2. Initialize Express App
const app = express();
const PORT = 5000;

// --- Global variable to store file paths for batching ---
let fileQueue = [];

// 3. Middleware Setup
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use(express.static(path.join(__dirname, '..', 'frontend')));

// Define paths
const outputDir = path.join(__dirname, '..', 'output');
const excelFilePath = path.join(outputDir, 'registrations.xlsx');
const certificateDir = path.join(__dirname, '..', 'frontend', 'certificate');
const certificateTemplatePath = path.join(certificateDir, 'certificate.html');
const certificateCssPath = path.join(certificateDir, 'certificate.css');
const applicationFormDir = path.join(__dirname, '..', 'frontend', 'application_form');
const applicationFormTemplatePath = path.join(applicationFormDir, 'index.html');
const applicationFormCssPath = path.join(applicationFormDir, 'style.css');

fs.mkdir(outputDir, { recursive: true });

// --- HELPER FUNCTIONS ---
async function updateExcelSheet(data) {
    const filePath = excelFilePath;
    let existingData = [];
    const columns = [
        { header: 'Admission No', key: 'admission_no', width: 15 },
        { header: 'SDMIS Ref No', key: 'sdmis_ref_no', width: 15 },
        { header: 'Course Name', key: 'courseName', width: 30 },
        { header: 'Department', key: 'department', width: 20 },
        { header: 'Duration', key: 'duration', width: 20 },
        { header: 'Applicant Name', key: 'applicantName', width: 30 },
        { header: 'DOB', key: 'dob', width: 15 },
        { header: 'Gender', key: 'gender', width: 10 },
        { header: 'Father Name', key: 'fatherName', width: 30 },
        { header: 'Mother Name', key: 'motherName', width: 30 },
        { header: 'Address', key: 'address', width: 40 },
        { header: 'Email', key: 'email', width: 30 },
        { header: 'Mobile', key: 'mobile', width: 20 },
        { header: 'Aadhar', key: 'aadhar', width: 20 },
        { header: 'Start Date', key: 'fromDate', width: 15 },
        { header: 'End Date', key: 'toDate', width: 15 },
        { header: 'Submission Date', key: 'submissionDate', width: 20 },
        { header: 'Caste', key: 'casteCategory', width: 15 },
        { header: 'Course Fees', key: 'course_fees', width: 15 },
        { header: 'Edu 1 Course', key: 'edu_course_1', width: 20 },
        { header: 'Edu 1 School', key: 'edu_school_1', width: 30 },
        { header: 'Edu 1 Spec', key: 'edu_spec_1', width: 20 },
        { header: 'Edu 1 Year', key: 'edu_year_1', width: 15 },
        { header: 'Edu 1 Perc', key: 'edu_perc_1', width: 15 },
        { header: 'Edu 2 Course', key: 'edu_course_2', width: 20 },
        { header: 'Edu 2 School', key: 'edu_school_2', width: 30 },
        { header: 'Edu 2 Spec', key: 'edu_spec_2', width: 20 },
        { header: 'Edu 2 Year', key: 'edu_year_2', width: 15 },
        { header: 'Edu 2 Perc', key: 'edu_perc_2', width: 15 },
    ];
    try {
        const tempWorkbook = new ExcelJS.Workbook();
        await tempWorkbook.xlsx.readFile(filePath);
        const worksheet = tempWorkbook.getWorksheet('Registrations');
        if (worksheet) {
            worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                if (rowNumber > 1) {
                    let rowData = {};
                    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                        if (columns[colNumber - 1]) {
                            rowData[columns[colNumber - 1].key] = cell.value;
                        }
                    });
                    existingData.push(rowData);
                }
            });
        }
    } catch (error) { /* File doesn't exist */ }
    
    const flatData = {
        ...data,
        edu_course_1: data.education?.[0]?.course,
        edu_school_1: data.education?.[0]?.school,
        edu_spec_1: data.education?.[0]?.spec,
        edu_year_1: data.education?.[0]?.year,
        edu_perc_1: data.education?.[0]?.perc,
        edu_course_2: data.education?.[1]?.course,
        edu_school_2: data.education?.[1]?.school,
        edu_spec_2: data.education?.[1]?.spec,
        edu_year_2: data.education?.[1]?.year,
        edu_perc_2: data.education?.[1]?.perc,
        submissionDate: new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' })
    };
    delete flatData.education;

    existingData.push(flatData);

    const newWorkbook = new ExcelJS.Workbook();
    const newWorksheet = newWorkbook.addWorksheet('Registrations');
    newWorksheet.columns = columns;
    newWorksheet.addRows(existingData);
    try {
        await newWorkbook.xlsx.writeFile(filePath);
        console.log('✅ Excel sheet overwritten successfully.');
    } catch (writeError) {
        console.error('❌ [Excel] ERROR: Failed to write to Excel file.', writeError);
    }
}

async function generateCertificate(data) {
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();
    let htmlTemplate = await fs.readFile(certificateTemplatePath, 'utf8');
    const cssContent = await fs.readFile(certificateCssPath, 'utf8');
    const msmeLogoPath = path.join(certificateDir, 'msme.jpeg');
    const citdLogoPath = path.join(certificateDir, 'citd main.png');
    const msmeLogoBase64 = await fs.readFile(msmeLogoPath, 'base64');
    const citdLogoBase64 = await fs.readFile(citdLogoPath, 'base64');
    const msmeDataUrl = `data:image/jpeg;base64,${msmeLogoBase64}`;
    const citdDataUrl = `data:image/png;base64,${citdLogoBase64}`;
    htmlTemplate = htmlTemplate
        .replace('src="msme.jpeg"', `src="${msmeDataUrl}"`)
        .replace('src="citd main.png"', `src="${citdDataUrl}"`);
    const finalHtml = htmlTemplate.replace('</head>', `<style>${cssContent}</style></head>`);
    await page.setContent(finalHtml, { waitUntil: 'networkidle0' });
    await page.evaluate(data => {
        const today = new Date();
        const issueDate = `${String(today.getDate()).padStart(2, '0')}.${String(today.getMonth() + 1).padStart(2, '0')}.${today.getFullYear()}`;
        document.getElementById('cert-gender').textContent = data.gender || '';
        document.getElementById('cert-name').textContent = ` ${data.applicantName.toUpperCase()}`;
        document.getElementById('cert-father').textContent = ` ${data.fatherName.toUpperCase()}`;
        document.getElementById('cert-course').textContent = `INTERNSHIP PROGRAMME ON ${data.courseName.toUpperCase()}`;
        document.getElementById('cert-start').textContent = data.fromDate;
        document.getElementById('cert-end').textContent = data.toDate;
        document.getElementById('cert-issue-date').textContent = issueDate;
        if (data.photo) {
            const photoElem = document.getElementById('cert-photo');
            if (photoElem) { photoElem.src = data.photo; }
        }
    }, data);
    const pdfFileName = `Certificate-${data.applicantName.replace(/\s+/g, '_')}-${Date.now()}.pdf`;
    const pdfPath = path.join(outputDir, pdfFileName);
    await page.pdf({ path: pdfPath, width: '1058px', height: '748px', printBackground: true });
    await browser.close();
    console.log(`✅ PDF certificate generated: ${pdfPath}`);
    return pdfPath;
}

async function generateApplicationFormPdf(data) {
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();
    let htmlTemplate = await fs.readFile(applicationFormTemplatePath, 'utf8');
    const cssContent = await fs.readFile(applicationFormCssPath, 'utf8');
    const msmeLogoPath = path.join(applicationFormDir, 'msme.jpeg');
    const citdLogoPath = path.join(applicationFormDir, 'citd main.png');
    const msmeLogoBase64 = await fs.readFile(msmeLogoPath, 'base64');
    const citdLogoBase64 = await fs.readFile(citdLogoPath, 'base64');
    const msmeDataUrl = `data:image/jpeg;base64,${msmeLogoBase64}`;
    const citdDataUrl = `data:image/png;base64,${citdLogoBase64}`;
    htmlTemplate = htmlTemplate
        .replace('src="msme.jpeg"', `src="${msmeDataUrl}"`)
        .replace('src="citd main.png"', `src="${citdDataUrl}"`);
    const finalHtml = htmlTemplate.replace('</head>', `<style>${cssContent}</style></head>`);
    await page.setContent(finalHtml, { waitUntil: 'networkidle0' });
    await page.evaluate(data => {
        document.getElementById('admission_no').value = data.admission_no || '';
        document.getElementById('sdmis_ref_no').value = data.sdmis_ref_no || '';
        document.getElementById('course-name').value = data.courseName || '';
        document.getElementById('department').value = data.department || '';
        document.getElementById('duration').value = data.duration || '';
        document.getElementById('from-date').value = data.fromDate || '';
        document.getElementById('to-date').value = data.toDate || '';
        document.getElementById('applicant-name').value = data.applicantName || '';
        document.getElementById('dob').value = data.dob || '';
        document.getElementById('father-name').value = data.fatherName || '';
        document.getElementById('mother-name').value = data.motherName || '';
        document.getElementById('address').value = data.address || '';
        document.getElementById('mobile').value = data.mobile || '';
        document.getElementById('email').value = data.email || '';
        document.getElementById('aadhar').value = data.aadhar || '';
        document.getElementById('course_fees').value = data.course_fees || '';
        if (data.casteCategory) {
            const categories = data.casteCategory.split(', ');
            categories.forEach(cat => {
                const checkbox = document.querySelector(`input[name="caste"][value="${cat}"]`);
                if (checkbox) checkbox.checked = true;
            });
        }
        if (data.education && data.education.length > 0) {
            document.querySelector('[name="edu_course_1"]').value = data.education[0].course || '';
            document.querySelector('[name="edu_school_1"]').value = data.education[0].school || '';
            document.querySelector('[name="edu_spec_1"]').value = data.education[0].spec || '';
            document.querySelector('[name="edu_year_1"]').value = data.education[0].year || '';
            document.querySelector('[name="edu_perc_1"]').value = data.education[0].perc || '';
        }
        if (data.education && data.education.length > 1) {
            document.querySelector('[name="edu_course_2"]').value = data.education[1].course || '';
            document.querySelector('[name="edu_school_2"]').value = data.education[1].school || '';
            document.querySelector('[name="edu_spec_2"]').value = data.education[1].spec || '';
            document.querySelector('[name="edu_year_2"]').value = data.education[1].year || '';
            document.querySelector('[name="edu_perc_2"]').value = data.education[1].perc || '';
        }
        if (data.gender === 'Mr.') {
            document.getElementById('gender_male').checked = true;
        } else if (data.gender === 'Ms.') {
            document.getElementById('gender_female').checked = true;
        }
        if (data.photo) {
            const photoElem = document.getElementById('preview');
            if (photoElem) {
                photoElem.src = data.photo;
                photoElem.style.display = 'block';
            }
        }
        const submitBtn = document.querySelector('.submit-btn');
        if(submitBtn) submitBtn.style.display = 'none';
    }, data);
    const pdfFileName = `Application-${data.applicantName.replace(/\s+/g, '_')}-${Date.now()}.pdf`;
    const pdfPath = path.join(outputDir, pdfFileName);
    await page.pdf({ path: pdfPath, format: 'A4', printBackground: true });
    await browser.close();
    console.log(`✅ PDF application form generated: ${pdfPath}`);
    return pdfPath;
}

// --- EMAIL FUNCTIONS ---
async function sendStudentEmail(data, applicationFormPath) {
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: { user: process.env.EMAIL_USER, pass: process.env.EMAIL_PASS },
    });
    const mailOptions = {
        from: `"CITD Hyderabad" <${process.env.EMAIL_USER}>`,
        to: data.email,
        subject: 'Application Received - CITD Short Term Course',
        html: `<h3>Dear ${data.applicantName},</h3><p>Thank you for registering...</p>`,
        attachments: [{
            filename: path.basename(applicationFormPath),
            path: applicationFormPath,
        }],
    };
    try {
        await transporter.sendMail(mailOptions);
        console.log(`✅ Confirmation email sent to student: ${data.email}`);
    } catch (error) {
        console.error(`❌ Failed to send email to student: ${data.email}`, error);
    }
}

async function sendBatchedFacultyEmail() {
    if (fileQueue.length === 0) {
        console.log('📧 No new registrations to send in this batch.');
        return;
    }
    console.log(`📧 Preparing to send batch email with ${fileQueue.length / 2} student(s) to faculty...`);
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: { user: process.env.EMAIL_USER, pass: process.env.EMAIL_PASS },
    });
    const attachments = fileQueue.map(item => ({
        filename: path.basename(item.path),
        path: item.path,
    }));
    attachments.push({
        filename: 'registrations.xlsx',
        path: excelFilePath,
    });
    const mailOptions = {
        from: `"CITD Registration System" <${process.env.EMAIL_USER}>`,
        to: process.env.FACULTY_EMAIL,
        subject: `Student Registration Batch Report - ${new Date().toLocaleTimeString('en-IN', { timeZone: 'Asia/Kolkata' })}`,
        html: `<h3>Batch Registration Report</h3><p>Please find attached documents for <strong>${fileQueue.length / 2}</strong> new student(s) who registered in this period.</p><p>The updated master registration list is also attached.</p>`,
        attachments: attachments,
    };
    try {
        await transporter.sendMail(mailOptions);
        console.log('✅ Batch email sent successfully to faculty.');
        fileQueue = []; // Clear the queue
    } catch (error) {
        console.error('❌ CRITICAL ERROR: Failed to send batch email to faculty.', error);
    }
}

// --- SCHEDULER ---
// CORRECTED: Runs ONCE at the specified time.
// Batch 1: Runs at 11:42 AM
cron.schedule('42 10 * * *', sendBatchedFacultyEmail, {
    timezone: "Asia/Kolkata"
});

// Batch 2: Runs at 11:45 AM
cron.schedule('44 10 * * *', sendBatchedFacultyEmail, {
    timezone: "Asia/Kolkata"
});

console.log('🕒 Email scheduler is running. Batches will be sent at 10:42 AM and 10:44 AM.');

// --- API ENDPOINT ---
app.post('/api/submit-form', async (req, res) => {
    console.log('\n-----------------------------------------');
    console.log(`➡️ Received new form submission at ${new Date().toLocaleTimeString()}`);
    try {
        const formData = req.body;
        const certificatePath = await generateCertificate(formData);
        const applicationFormPath = await generateApplicationFormPdf(formData);
        await updateExcelSheet(formData);
        await sendStudentEmail(formData, applicationFormPath);
        fileQueue.push({ type: 'certificate', path: certificatePath });
        fileQueue.push({ type: 'application', path: applicationFormPath });
        console.log(`📥 Files for ${formData.applicantName} added to the faculty email queue. Current queue size: ${fileQueue.length} files.`);
        console.log('🎉 All tasks completed successfully!');
        res.status(200).json({ message: 'Registration successful! You will receive a confirmation email shortly.' });
    } catch (error) {
        console.error('❌ An error occurred during processing:', error);
        res.status(500).json({ message: 'An error occurred on the server.', error: error.message });
    } finally {
        console.log('-----------------------------------------\n');
    }
});

app.listen(PORT, () => {
    console.log(`✅ Server running on http://localhost:${PORT}`);
    console.log(`➡️ Access your form at: http://localhost:${PORT}/application_form/index.html`);
});

