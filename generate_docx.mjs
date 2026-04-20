import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  BorderStyle, TabStopPosition, TabStopType, ShadingType,
  Table, TableRow, TableCell, WidthType, TableBorders, ImageRun
} from 'docx';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ── Color constants ──
const ACCENT = "0EA5E9";
const DARK = "1E293B";
const MUTED = "64748B";
const LIGHT_BG = "F0F9FF";

// ── Helper functions ──
function heading(text, level = HeadingLevel.HEADING_1) {
  return new Paragraph({
    heading: level,
    spacing: { before: 300, after: 100 },
    children: [
      new TextRun({ text, bold: true, size: level === HeadingLevel.HEADING_1 ? 28 : 22, color: DARK, font: "Calibri" })
    ],
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 2, color: ACCENT }
    }
  });
}

function bullet(text, options = {}) {
  const runs = [];
  // Parse bold markers **text**
  const parts = text.split(/(\*\*[^*]+\*\*)/g);
  for (const part of parts) {
    if (part.startsWith('**') && part.endsWith('**')) {
      runs.push(new TextRun({ text: part.slice(2, -2), bold: true, size: 20, color: DARK, font: "Calibri" }));
    } else {
      runs.push(new TextRun({ text: part, size: 20, color: MUTED, font: "Calibri" }));
    }
  }
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    bullet: { level: options.level || 0 },
    children: runs,
  });
}

function spacer(pts = 100) {
  return new Paragraph({ spacing: { before: pts } });
}

// ── Document creation ──
const doc = new Document({
  creator: "Emmanuel Gonzalez L.",
  title: "Emmanuel Gonzalez L. — QA Automation Engineer & SDET",
  description: "Professional CV of Emmanuel Gonzalez L., Senior QA Automation Engineer",
  styles: {
    default: {
      document: {
        run: { font: "Calibri", size: 20, color: DARK }
      }
    }
  },
  sections: [{
    properties: {
      page: {
        margin: { top: 720, bottom: 720, left: 900, right: 900 }
      }
    },
    children: [
      // ═══════════════════════════════════════
      // HEADER
      // ═══════════════════════════════════════
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
        children: [
          new TextRun({ text: "EMMANUEL GONZALEZ L.", bold: true, size: 36, color: DARK, font: "Calibri" })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 40 },
        children: [
          new TextRun({ text: "SENIOR QA AUTOMATION ENGINEER  ·  SDET  ·  QUALITY ARCHITECT", bold: true, size: 20, color: ACCENT, font: "Calibri" })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [
          new TextRun({ text: "📞 +52 999 338 0373  ·  ✉ emanuel2464@gmail.com  ·  📍 Progreso, Yucatán, México", size: 18, color: MUTED, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [
          new TextRun({ text: "🔗 github.com/Alucardo18  ·  💼 linkedin.com/in/emmanuel-gonzalez-lopez-6614ba58", size: 18, color: MUTED, font: "Calibri" }),
        ]
      }),

      // ═══════════════════════════════════════
      // PROFESSIONAL SUMMARY
      // ═══════════════════════════════════════
      heading("PROFESSIONAL SUMMARY"),
      new Paragraph({
        spacing: { before: 100, after: 200 },
        shading: { type: ShadingType.SOLID, color: LIGHT_BG },
        children: [
          new TextRun({ text: "Result-driven ", size: 20, color: MUTED, font: "Calibri" }),
          new TextRun({ text: "Senior QA Automation Engineer & SDET", bold: true, size: 20, color: DARK, font: "Calibri" }),
          new TextRun({ text: " with ", size: 20, color: MUTED, font: "Calibri" }),
          new TextRun({ text: "10+ years", bold: true, size: 20, color: ACCENT, font: "Calibri" }),
          new TextRun({ text: " of experience delivering high-quality software across ", size: 20, color: MUTED, font: "Calibri" }),
          new TextRun({ text: "large-scale payment platforms and financial systems", bold: true, size: 20, color: DARK, font: "Calibri" }),
          new TextRun({ text: ". Proven track record in architecting robust BDD automation frameworks using Java, Cucumber, and Selenium, while integrating tests into CI/CD pipelines. Currently pioneering ", size: 20, color: MUTED, font: "Calibri" }),
          new TextRun({ text: "AI-augmented quality engineering", bold: true, size: 20, color: DARK, font: "Calibri" }),
          new TextRun({ text: " workflows using Cursor, Claude, Codex, and Antigravity to accelerate test development and defect prediction. Passionate about shifting quality left and driving engineering excellence across cross-functional teams.", size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),

      // ═══════════════════════════════════════
      // KEY ACHIEVEMENTS
      // ═══════════════════════════════════════
      heading("KEY ACHIEVEMENTS"),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: TableBorders.NONE,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 25, type: WidthType.PERCENTAGE },
                shading: { type: ShadingType.SOLID, color: LIGHT_BG },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "10+", bold: true, size: 28, color: ACCENT, font: "Calibri" })] }),
                  new Paragraph({ children: [new TextRun({ text: "Years of QA & Automation expertise", size: 16, color: MUTED, font: "Calibri" })] }),
                ]
              }),
              new TableCell({
                width: { size: 25, type: WidthType.PERCENTAGE },
                shading: { type: ShadingType.SOLID, color: LIGHT_BG },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "US Visa", bold: true, size: 28, color: ACCENT, font: "Calibri" })] }),
                  new Paragraph({ children: [new TextRun({ text: "Active B1/B2 Tourist Visa", size: 16, color: MUTED, font: "Calibri" })] }),
                ]
              }),
              new TableCell({
                width: { size: 25, type: WidthType.PERCENTAGE },
                shading: { type: ShadingType.SOLID, color: LIGHT_BG },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "AI-First", bold: true, size: 28, color: ACCENT, font: "Calibri" })] }),
                  new Paragraph({ children: [new TextRun({ text: "Agentic programming pioneer", size: 16, color: MUTED, font: "Calibri" })] }),
                ]
              }),
              new TableCell({
                width: { size: 25, type: WidthType.PERCENTAGE },
                shading: { type: ShadingType.SOLID, color: LIGHT_BG },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "E2E", bold: true, size: 28, color: ACCENT, font: "Calibri" })] }),
                  new Paragraph({ children: [new TextRun({ text: "API, UI, Integration & Performance", size: 16, color: MUTED, font: "Calibri" })] }),
                ]
              }),
            ]
          }),
        ]
      }),
      spacer(200),

      // ═══════════════════════════════════════
      // PROFESSIONAL EXPERIENCE
      // ═══════════════════════════════════════
      heading("PROFESSIONAL EXPERIENCE"),

      // --- Current Experience ---
      new Paragraph({
        spacing: { before: 200, after: 40 },
        children: [
          new TextRun({ text: "Senior QA Automation Engineer", bold: true, size: 24, color: DARK, font: "Calibri" }),
          new TextRun({ text: "    ", size: 20 }),
          new TextRun({ text: "2019 — Present", bold: true, size: 18, color: ACCENT, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        spacing: { after: 80 },
        children: [
          new TextRun({ text: "Freelance / Independent Consultant", bold: true, size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),
      bullet("Lead QA automation strategy for large-scale software systems, **ensuring quality gates** across critical services."),
      bullet("Architected and maintained **BDD automation frameworks** using Java, Cucumber, and Selenium, enabling continuous regression testing within CI/CD pipelines managed via **Jenkins**."),
      bullet("Designed and executed comprehensive **REST API test suites** for services consuming legacy systems, validating transaction flows."),
      bullet("Integrated **AI-driven tools (Cursor, Claude, Codex, Antigravity)** into daily automation workflows, pioneering agentic programming techniques to accelerate test case generation and framework development."),
      bullet("Managed test data orchestration leveraging **MongoDB** and validated event-driven architectures through **Kafka** message streams."),
      bullet("Performed **load & performance testing** using JMeter to ensure system resilience under peak transaction volumes."),
      bullet("Collaborated with cross-functional engineering teams across multiple geographies using Agile/Scrum, contributing to sprint planning, code reviews, and quality retrospectives."),

      // --- 4th Source ---
      new Paragraph({
        spacing: { before: 300, after: 40 },
        children: [
          new TextRun({ text: "QA Automation Engineer", bold: true, size: 24, color: DARK, font: "Calibri" }),
          new TextRun({ text: "    ", size: 20 }),
          new TextRun({ text: "2015 — 2019", bold: true, size: 18, color: ACCENT, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        spacing: { after: 80 },
        children: [
          new TextRun({ text: "4th Source — Enterprise Software Consulting", bold: true, size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),
      bullet("Created and executed comprehensive test plans covering **functional, non-functional, regression, and integration** testing scenarios."),
      bullet("Built API automation suites using **REST template from Spring Boot**, testing services secured with **OAuth 1.0** authentication protocols."),
      bullet("Implemented **Behavior-Driven Development (BDD)** test frameworks using Cucumber with Java, improving test readability and stakeholder communication."),
      bullet("Automated landing and login page flows using **Selenium WebDriver** with Java, reducing manual regression effort."),
      bullet("Deployed and managed builds through **CI/CD pipelines** using Jenkins, Maven, and JFrog Artifactory with code versioned in GitHub."),
      bullet("Worked within **Agile methodologies (Scrum/Kanban)** using Jira and Confluence for sprint tracking and knowledge management."),
      bullet("Operated in cloud-native environments using **Pivotal Cloud Foundry (PCF)** and event-driven architectures with **Kafka**."),

      // --- CENACE ---
      new Paragraph({
        spacing: { before: 300, after: 40 },
        children: [
          new TextRun({ text: "Technical Support Engineer", bold: true, size: 24, color: DARK, font: "Calibri" }),
          new TextRun({ text: "    ", size: 20 }),
          new TextRun({ text: "2014 — 2015", bold: true, size: 18, color: ACCENT, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        spacing: { after: 80 },
        children: [
          new TextRun({ text: "CENACE — Centro Nacional de Control de Energía", bold: true, size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),
      bullet("Managed Windows Server infrastructure **(W2008-R2 & W2012-R2)** supporting critical national energy control operations."),
      bullet("Configured and deployed **CISCO networking equipment** (Catalyst 6500 & 2970s), PoE telephones, and wireless access points."),
      bullet("Developed a **web-based autonomous monitoring and temperature control system** for data center facilities, combining hardware and software engineering."),
      bullet("Provided tier-2 technical support and managed service desk operations for routing and infrastructure inquiries."),

      // ═══════════════════════════════════════
      // TECHNICAL SKILLS
      // ═══════════════════════════════════════
      heading("TECHNICAL SKILLS"),
      new Paragraph({
        spacing: { before: 100, after: 60 },
        children: [
          new TextRun({ text: "Languages & Core:  ", bold: true, size: 20, color: DARK, font: "Calibri" }),
          new TextRun({ text: "Java  ·  SQL  ·  C++  ·  Visual Basic", size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({ text: "Automation & Testing:  ", bold: true, size: 20, color: DARK, font: "Calibri" }),
          new TextRun({ text: "Selenium  ·  Cucumber / BDD  ·  REST API Testing  ·  Postman  ·  SOAPUI  ·  JMeter  ·  Spring Boot", size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({ text: "AI & Emerging Tech:  ", bold: true, size: 20, color: DARK, font: "Calibri" }),
          new TextRun({ text: "Agentic Programming  ·  Cursor  ·  Claude  ·  Codex  ·  Antigravity", size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({ text: "DevOps & Tools:  ", bold: true, size: 20, color: DARK, font: "Calibri" }),
          new TextRun({ text: "GitHub  ·  Jenkins  ·  Maven  ·  JFrog Artifactory  ·  Jira  ·  Confluence  ·  IntelliJ IDEA", size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({ text: "Infrastructure & Data:  ", bold: true, size: 20, color: DARK, font: "Calibri" }),
          new TextRun({ text: "MongoDB  ·  Kafka  ·  PCF (Pivotal Cloud Foundry)", size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        spacing: { after: 100 },
        children: [
          new TextRun({ text: "Methodologies:  ", bold: true, size: 20, color: DARK, font: "Calibri" }),
          new TextRun({ text: "Agile / Scrum  ·  Kanban  ·  BDD / TDD  ·  Shift-Left Testing  ·  CI/CD  ·  Waterfall", size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),

      // ═══════════════════════════════════════
      // PROJECTS & INNOVATION
      // ═══════════════════════════════════════
      heading("PROJECTS & INNOVATION"),
      new Paragraph({
        spacing: { before: 120, after: 40 },
        children: [
          new TextRun({ text: "📱 IntelicaLabs Controller", bold: true, size: 22, color: DARK, font: "Calibri" }),
          new TextRun({ text: " — Published Android App", size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({ text: "Self-taught Android development; designed, built, and published a Bluetooth-enabled robotics controller app on Google Play Store. Demonstrates full software lifecycle ownership from concept to deployment.", size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: "Tech: ", bold: true, size: 18, color: DARK, font: "Calibri" }),
          new TextRun({ text: "Android  ·  Java  ·  Bluetooth  ·  Published App", size: 18, color: ACCENT, font: "Calibri" }),
        ]
      }),

      new Paragraph({
        spacing: { before: 160, after: 40 },
        children: [
          new TextRun({ text: "🤖 Competitive Robotics", bold: true, size: 22, color: DARK, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({ text: "Designed and programmed autonomous robots for regional competitions, combining electronics engineering with embedded software.", size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),
      bullet("🥇 **1st Place** — 9th Robotics Contest of Mayab (Line Follower Autonomous Robot)"),
      bullet("🏅 **4th Place** — 9th Robotics Contest of Mayab (MiniSumo Autonomous Robot)"),
      bullet("🥇 **1st Place** — 8th Robotics Contest of Mayab (Line Follower Autonomous Robot)"),

      // ═══════════════════════════════════════
      // EDUCATION
      // ═══════════════════════════════════════
      heading("EDUCATION"),
      new Paragraph({
        spacing: { before: 120, after: 20 },
        children: [
          new TextRun({ text: "B.Sc. Electronics Engineering", bold: true, size: 22, color: DARK, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        spacing: { after: 20 },
        children: [
          new TextRun({ text: "Instituto Tecnológico de Mérida — Mérida, Yucatán, México", size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        spacing: { after: 100 },
        children: [
          new TextRun({ text: "2009 — 2014", bold: true, size: 18, color: ACCENT, font: "Calibri" }),
        ]
      }),

      // ═══════════════════════════════════════
      // LANGUAGES
      // ═══════════════════════════════════════
      heading("LANGUAGES"),
      new Paragraph({
        spacing: { before: 100, after: 40 },
        children: [
          new TextRun({ text: "Spanish ", bold: true, size: 20, color: DARK, font: "Calibri" }),
          new TextRun({ text: "— Native", size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: "English ", bold: true, size: 20, color: DARK, font: "Calibri" }),
          new TextRun({ text: "— Professional Working Proficiency (85%)", size: 20, color: MUTED, font: "Calibri" }),
        ]
      }),
    ]
  }]
});

// ── Generate DOCX ──
const outputPath = path.join(__dirname, 'Emmanuel_Gonzalez_Antigravity_CV.docx');
const buffer = await Packer.toBuffer(doc);
fs.writeFileSync(outputPath, buffer);
console.log(`✅ DOCX saved to: ${outputPath}`);
console.log('📝 Upload this file to Google Drive, then open with Google Docs to edit!');
