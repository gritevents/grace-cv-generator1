const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, BorderStyle, WidthType, ShadingType,
  TabStopType
} = require('docx');
const fs = require('fs');

const GREEN = "1F7A4D";
const GREEN_LIGHT = "E8F5EE";
const BLACK = "000000";
const GRAY = "444444";
const MID_GRAY = "666666";

function sectionHeader(text) {
  return new Paragraph({
    spacing: { before: 280, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: GREEN, space: 4 } },
    children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 22, color: GREEN, font: "Arial" })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 20, font: "Arial", color: BLACK })]
  });
}

function jobBlock(company, location, dates, title) {
  return [
    new Paragraph({
      spacing: { before: 200, after: 20 },
      tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
      children: [
        new TextRun({ text: company, bold: true, size: 22, color: BLACK, font: "Arial" }),
        new TextRun({ text: location ? ` — ${location}` : "", size: 20, color: GRAY, font: "Arial" }),
        new TextRun({ text: "\t", size: 20, font: "Arial" }),
        new TextRun({ text: dates, size: 18, color: MID_GRAY, font: "Arial", italics: true }),
      ]
    }),
    new Paragraph({
      spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: title, bold: true, size: 20, color: GREEN, font: "Arial", italics: true })]
    })
  ];
}

function skillCheckRow(items) {
  const cellW = 3120;
  return new TableRow({
    children: items.map(item => new TableCell({
      width: { size: cellW, type: WidthType.DXA },
      borders: {
        top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
        left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }
      },
      shading: { fill: GREEN_LIGHT, type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [new Paragraph({
        children: [
          new TextRun({ text: "✓  ", size: 18, color: GREEN, font: "Arial", bold: true }),
          new TextRun({ text: item, size: 18, color: BLACK, font: "Arial" })
        ]
      })]
    }))
  });
}

const skillRows = [
  ["Retail & Floor Sales", "Customer Needs Assessment", "Product Demonstration"],
  ["Closing & Converting", "Upselling & Cross-Selling", "Objection Handling"],
  ["Brand Activations & Promotions", "Field Sales", "Target Achievement"],
  ["Appointment Setting", "Outbound Prospecting", "Corporate Sales & B2B"],
  ["Pipeline & CRM Management", "Follow-Up Sequences", "Client Retention"],
  ["Multi-Channel Outreach", "WhatsApp & Social Selling", "Confident English Communicator"],
];

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "•",
        alignment: AlignmentType.LEFT,
        style: {
          paragraph: { indent: { left: 440, hanging: 280 } },
          run: { color: GREEN, size: 20, font: "Arial" }
        }
      }]
    }]
  },
  styles: { default: { document: { run: { font: "Arial", size: 20, color: BLACK } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 900, right: 1080, bottom: 900, left: 1080 }
      }
    },
    children: [

      // ── NAME & HEADLINE ──
      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: "GRACE W. NDUNGU", bold: true, size: 56, color: GREEN, font: "Arial" })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: "Sales Professional  •  Retail & Brand Sales  •  Client Acquisition  •  Business Development", size: 20, color: GRAY, font: "Arial" })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 200 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GREEN, space: 4 } },
        children: [new TextRun({ text: "+254 707 768 071  |  +254 701 441 045  |  gracetitusn@gmail.com  |  Kiambu, Nairobi  |  linkedin.com/in/grace-ndungu-903179b8", size: 18, color: GRAY, font: "Arial" })]
      }),

      // ── PROFESSIONAL SUMMARY ──
      sectionHeader("Professional Summary"),
      new Paragraph({
        spacing: { before: 80, after: 140 },
        children: [new TextRun({
          text: "High-energy, results-driven sales professional with 10+ years of experience in customer-facing sales, client acquisition, business development, and brand promotion. Personally prospected, pitched, and closed multiple high-value corporate event contracts and retail studio clients at Kajim Productions — converting cold conversations into confirmed, paying, repeat business. Passionate about Proudly Kenyan brands and committed to driving revenue through genuine relationships, confident pitching, and consistent follow-through. Experienced in retail sales, field activations, corporate B2B sales, and multi-channel outreach. A natural communicator who thrives in target-driven environments and never backs down from a challenge. Fluent in English and Kiswahili. Trained in Public Speaking by renowned orator Benny Hinn Walubengo (iSpeak Academy, 2024).",
          size: 20, font: "Arial", color: BLACK
        })]
      }),

      // ── KEY SALES SKILLS TABLE ──
      sectionHeader("Key Sales Skills & Tools"),
      new Paragraph({ spacing: { before: 80, after: 60 }, children: [] }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 3120, 3120],
        rows: skillRows.map(row => skillCheckRow(row))
      }),

      // ── SALES ACHIEVEMENTS ──
      sectionHeader("Sales Achievements Snapshot"),
      bullet("Personally prospected, pitched, and closed multiple podcast studio bookings at Kajim Productions — converting cold conversations into confirmed, paid, repeat clients through confident pitching and relationship building."),
      bullet("Networked with multiple high-end corporate clients at Kajim Productions, delivered compelling presentations of full event production services, and converted those conversations into confirmed, paying contracts — demonstrating strong B2B and corporate sales ability."),
      bullet("Managed Kajim Productions' Facebook and Instagram pages — creating compelling sales-driven content that attracted new enquiries and supported brand growth and client acquisition."),
      bullet("Conducted a full Instagram account audit for Vista Digital — reviewed and corrected 500+ posts to fix location data, directly improving Google My Business visibility and local customer discoverability."),
      bullet("Led outbound outreach campaigns at Women Tech Circles that converted leads into enrolled, paying community members — using targeted scripts and Zoho CRM follow-up sequences to improve conversion rates."),
      bullet("Managed full appointment pipelines for multiple remote clients as a Freelance Sales VA — qualifying leads, booking sales calls, and driving revenue through consistent, strategic follow-up."),
      bullet("Maintained 100% CEO appointment show-up rate at Kajim Productions through strategic scheduling and follow-up — directly protecting revenue opportunities."),
      bullet("Promoted at SamaSource Kenya from Agent to Reviewer based on consistent target achievement and quality output — a clear demonstration of the performance discipline that defines top sales professionals."),

      // ── WHY DENRI AFRICA ──
      sectionHeader("Why Denri Africa"),
      new Paragraph({
        spacing: { before: 80, after: 140 },
        children: [new TextRun({
          text: "Denri Africa is not just a bag brand — it is a movement. A Proudly Kenyan success story that started with one store and grew to 19+ outlets nationwide by building a product that Kenyans are genuinely proud to carry. That is the kind of brand I want to sell for. I am passionate about locally-made products, I believe in Denri's quality and vision, and I have the energy, communication skills, and sales drive to represent the brand with the excellence it deserves — in-store, in the field, and in every customer interaction.",
          size: 20, font: "Arial", color: BLACK, italics: true
        })]
      }),

      // ── PROFESSIONAL EXPERIENCE ──
      sectionHeader("Professional Experience"),

      ...jobBlock("Kajim Events Ltd / Kajim Productions", "Nairobi, Kenya", "June 2025 – March 2026", "Client Relations, Sales & Business Development Lead"),
      bullet("Personally prospected and pitched multiple clients for Kajim Productions' podcast studio — explained the product offering, handled objections confidently, and closed multiple paid bookings, with several clients returning for repeat sessions."),
      bullet("Networked with multiple high-end corporate clients, delivered compelling pitches on Kajim's full event production services, and successfully converted those conversations into confirmed, paying event contracts."),
      bullet("Managed Kajim Productions' Facebook and Instagram pages — creating and scheduling brand content, engaging with followers, and responding to enquiries to drive new business."),
      bullet("Served as the first point of contact for all inbound enquiries — qualifying leads, assessing fit, and booking confirmed meetings with the CEO, ensuring zero leads were lost."),
      bullet("Built lasting client relationships through warm, professional, and persuasive communication — consistently commended by the CEO, General Manager, and clients for sales performance and service excellence."),
      bullet("Demonstrated versatility in a lean startup environment: simultaneously managing PA duties, sales, client relations, and on-field videography — the hallmark of a driven, multi-skilled sales professional."),

      ...jobBlock("Vista Digital", "Nairobi, Kenya", "June 2025 – 2025", "Videographer, Photographer & Brand Social Media Specialist"),
      bullet("Produced compelling visual content for marketing and brand campaigns — directly supporting sales and product promotion goals."),
      bullet("Conducted a full audit and cleanup of Vista Digital's Instagram account — reviewed 500+ marketing posts, corrected location tags, and removed irrelevant content to strengthen brand presence."),
      bullet("Improved Google My Business local search visibility by updating historical post location data — directly benefiting the company's ability to attract walk-in and local customers."),

      ...jobBlock("Women Tech Circles (WTC)", "Nairobi, Kenya", "2023 – 2024", "Community Growth, Sales Outreach & Conversion Lead"),
      bullet("Led outbound outreach and sales conversion campaigns targeting prospective students and corporate partners — qualifying interest, following up persistently, and converting leads into enrolled, paying members."),
      bullet("Developed outreach scripts, email sequences, and CRM-driven sales strategies using Zoho CRM that improved conversion rates and community revenue."),
      bullet("Managed the full sales funnel — from first contact to payment confirmation — ensuring a seamless and persuasive experience for every prospect."),

      ...jobBlock("Self-Employed", "Remote", "2022 – Present", "Freelance Sales VA & Client Acquisition Manager"),
      bullet("Managed outbound and inbound sales communications for multiple clients — qualifying leads, scheduling sales calls, and managing full appointment pipelines."),
      bullet("Maintained consistent follow-up sequences via email, WhatsApp, and chat — moving prospects through the funnel and converting interest into booked, paying clients."),
      bullet("Used Zapier and Make to automate follow-up workflows — increasing response speed and improving lead-to-close ratios."),

      ...jobBlock("Jesus Is Alive Ministries (JIAM)", "Nairobi, Kenya", "2013 – 2019", "Customer Relations & Outreach Representative — 6 Years"),
      bullet("Delivered high-volume outbound and inbound communication for 6 years — managing appointment scheduling, follow-ups, and relationship management across phone, email, and in-person channels."),
      bullet("Conducted outreach campaigns — persuading, engaging, and converting contacts into active, committed participants through persistent and positive communication."),
      bullet("Demonstrated the patience, resilience, and warmth needed to engage diverse people and consistently drive action."),

      ...jobBlock("SamaSource Kenya", "Nairobi, Kenya", "2021 – 2022", "AI Associate — Promoted from Agent to Reviewer"),
      bullet("Consistently met and exceeded targets in a high-accountability, KPI-driven environment."),
      bullet("Promoted to Reviewer based on exceptional diligence and quality output — overseeing agent work and quality assurance before delivery to international clients."),

      // ── EDUCATION ──
      sectionHeader("Education & Certifications"),
      bullet("Diploma in Business Management — St. Paul's University (2016)."),
      bullet("Public Speaking Certification — iSpeak Academy, Orator Benny Hinn Walubengo (2024)."),
      bullet("Virtual Assistant Specialization — Coursera (2024)."),
      bullet("Google UX Design Professional Certificate — Coursera (2023)."),
      bullet("Software Engineering & Web Development — Power Learn Project (2023)."),
      bullet("Cybersecurity Awareness Training — SamaSource Kenya (2021)."),

      // ── LANGUAGES ──
      sectionHeader("Languages"),
      new Paragraph({
        spacing: { before: 80, after: 80 },
        children: [
          new TextRun({ text: "English", bold: true, size: 20, font: "Arial", color: BLACK }),
          new TextRun({ text: " — Fluent, Written & Spoken (Professional Level)     ", size: 20, font: "Arial", color: GRAY }),
          new TextRun({ text: "Kiswahili", bold: true, size: 20, font: "Arial", color: BLACK }),
          new TextRun({ text: " — Fluent, Written & Spoken.", size: 20, font: "Arial", color: GRAY }),
        ]
      }),

      // ── WHY I AM A GREAT SALES PROFESSIONAL ──
      sectionHeader("Why I Am a Great Sales Professional"),
      new Paragraph({
        spacing: { before: 80, after: 140 },
        children: [new TextRun({
          text: "I have spent over a decade doing exactly what sales requires — starting conversations, building trust fast, handling objections, and getting people to commit. I have closed real clients, brought in real revenue, and built real relationships from scratch. I am not afraid of rejection. I am energised by targets. I show up, follow through, and deliver results — every single time.",
          size: 20, font: "Arial", color: BLACK, italics: true
        })]
      }),

      // ── REFEREES ──
      sectionHeader("Referees"),
      new Paragraph({
        spacing: { before: 100, after: 20 },
        children: [new TextRun({ text: "McDonald Namu", bold: true, size: 20, font: "Arial", color: BLACK })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 20 },
        children: [new TextRun({ text: "Project Manager — SamaSource Kenya", size: 20, font: "Arial", color: GRAY, italics: true })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: "Phone: +254 743 373 400", size: 20, font: "Arial", color: BLACK })]
      }),
      new Paragraph({
        spacing: { before: 60, after: 20 },
        children: [new TextRun({ text: "MC Jimmie Kajim / James Wambui", bold: true, size: 20, font: "Arial", color: BLACK })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 20 },
        children: [new TextRun({ text: "Director — Kajim Events Limited / Kajim Productions", size: 20, font: "Arial", color: GRAY, italics: true })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: "Phone: +254 719 366 693", size: 20, font: "Arial", color: BLACK })]
      }),
      new Paragraph({
        spacing: { before: 60, after: 20 },
        children: [new TextRun({ text: "Bishop Margaret Wanjiru", bold: true, size: 20, font: "Arial", color: BLACK })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 20 },
        children: [new TextRun({ text: "Senior Leader — Jesus Is Alive Ministries (JIAM)", size: 20, font: "Arial", color: GRAY, italics: true })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: "Phone: +254 722 136 999", size: 20, font: "Arial", color: BLACK })]
      }),

      // ── MOTTO ──
      new Paragraph({
        spacing: { before: 200, after: 0 },
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: GREEN_LIGHT, space: 4 } },
        children: [new TextRun({ text: '"Excellence through innovation, service, and strategy."', italics: true, size: 18, color: MID_GRAY, font: "Arial" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/Grace_Ndungu_Denri_Sales_CV.docx', buf);
  console.log('Done');
});
