const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Add Excel CMS handling
const contentExcelPath = path.join(__dirname, 'content.xlsx');

// Initialize content.xlsx if it doesn't exist
if (!fs.existsSync(contentExcelPath)) {
    const wb = xlsx.utils.book_new();
    
    // Stats Sheet
    const wsStats = xlsx.utils.aoa_to_sheet([
        ["Metric", "Value", "Description"],
        ["GPA", "3.9", "Cumulative GPA"],
        ["Projects", "5+", "Major Projects"],
        ["Volunteering", "100h", "Volunteering"]
    ]);
    xlsx.utils.book_append_sheet(wb, wsStats, "Stats");

    // About Sheet
    const wsAbout = xlsx.utils.aoa_to_sheet([
        ["Key", "Text"],
        ["p1", "I am a highly motivated scholarship student pursuing a degree in Computer Science. My goal is to leverage my academic background to drive positive impact in the tech industry."],
        ["p2", "Through various leadership roles and academic excellence, I have proven my dedication to continuous learning and community engagement."]
    ]);
    xlsx.utils.book_append_sheet(wb, wsAbout, "About");

    // Timeline Sheet
    const wsTimeline = xlsx.utils.aoa_to_sheet([
        ["Section", "Title", "Date", "Description"],
        ["Academics", "Dean's Excellence Scholarship", "2023 - Present", "Awarded for outstanding academic performance and leadership potential among incoming freshmen."],
        ["Academics", "National Science Fair - 1st Place", "2022", "Developed an AI model predicting crop yields to assist local farmers, winning top honors nationally."],
        ["Experience", "Software Engineering Intern", "Summer 2023", "Assisted in the backend development of a scalable REST API using Node.js and improved database query performance by 20%."],
        ["Experience", "Past Community Development Leader", "2020 - 2022", "Spearheaded local neighborhood revitalization projects, organizing over 200 volunteers for after-school mentorship programs and community gardening."]
    ]);
    xlsx.utils.book_append_sheet(wb, wsTimeline, "Timeline");

    xlsx.writeFile(wb, contentExcelPath);
    console.log("Created default content.xlsx file successfully!");
}

// Endpoint to read Excel data
app.get('/api/content', (req, res) => {
    try {
        const wb = xlsx.readFile(contentExcelPath);
        const content = {};
        for (const sheetName of wb.SheetNames) {
            content[sheetName.toLowerCase()] = xlsx.utils.sheet_to_json(wb.Sheets[sheetName]);
        }
        res.status(200).json(content);
    } catch(err) {
        console.error("Error reading Excel:", err);
        res.status(500).json({ error: "Failed to read excel file" });
    }
});

// Serve static frontend files
app.use(express.static(path.join(__dirname, 'public')));

// API endpoint for contact form submission
app.post('/api/contact', (req, res) => {
    const { name, email, message } = req.body;
    
    // In a real application, you might save this to a database or send an email here.
    // For this portfolio, we'll just log it and send a success response.
    console.log('--- New Contact Form Submission ---');
    console.log(`Name: ${name}`);
    console.log(`Email: ${email}`);
    console.log(`Message: ${message}`);
    console.log('-----------------------------------');

    res.status(200).json({ success: true, message: 'Thank you for your message! I will get in touch soon.' });
});

// Fallback to serve index.html for any other route (SPA behavior)
app.use((req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
