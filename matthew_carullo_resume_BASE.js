const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  LevelFormat, BorderStyle, TabStopType, TabStopPosition,
  ExternalHyperlink
} = require('docx');
const fs = require('fs');

const BLUE = "1a3a5c";
const GRAY = "555555";
const BLACK = "111111";

function sectionHeader(text) {
  return new Paragraph({
    spacing: { before: 200, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLUE, space: 1 } },
    children: [
      new TextRun({
        text: text.toUpperCase(),
        bold: true,
        size: 22,
        color: BLUE,
        font: "Arial",
      })
    ]
  });
}

function jobHeader(title, company, location, dates) {
  return new Paragraph({
    spacing: { before: 150, after: 60 },
    tabStops: [{ type: TabStopType.RIGHT, position: 10080 }],
    children: [
      new TextRun({ text: title, bold: true, size: 22, font: "Arial", color: BLACK }),
      new TextRun({ text: "  |  ", size: 22, font: "Arial", color: GRAY }),
      new TextRun({ text: company, bold: true, size: 22, font: "Arial", color: BLUE }),
      new TextRun({ text: "\t", size: 22, font: "Arial" }),
      new TextRun({ text: `${location}  |  ${dates}`, size: 21, font: "Arial", color: GRAY, italics: true }),
    ]
  });
}

function bullet(text, bold_prefix = null) {
  const children = [];
  if (bold_prefix) {
    children.push(new TextRun({ text: bold_prefix + " ", bold: true, size: 21, font: "Arial", color: BLACK }));
    children.push(new TextRun({ text: text, size: 21, font: "Arial", color: BLACK }));
  } else {
    children.push(new TextRun({ text: text, size: 21, font: "Arial", color: BLACK }));
  }
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 55, after: 55 },
    children
  });
}

function skillRow(category, items) {
  return new Paragraph({
    spacing: { before: 55, after: 55 },
    children: [
      new TextRun({ text: category + ":  ", bold: true, size: 21, font: "Arial", color: BLACK }),
      new TextRun({ text: items, size: 21, font: "Arial", color: GRAY }),
    ]
  });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 440, hanging: 260 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 576, right: 864, bottom: 576, left: 864 }
      }
    },
    children: [

      // NAME
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 30 },
        children: [
          new TextRun({ text: "Matthew Carullo", bold: true, size: 52, font: "Arial", color: BLACK })
        ]
      }),

      // CONTACT
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 30 },
        children: [
          new TextRun({ text: "+1 (416) 580-9227  |  ", size: 18, font: "Arial", color: GRAY }),
          new ExternalHyperlink({
            link: "mailto:mcarullo@uwaterloo.ca",
            children: [new TextRun({ text: "mcarullo@uwaterloo.ca", style: "Hyperlink", size: 18, font: "Arial" })]
          }),
          new TextRun({ text: "  |  ", size: 18, font: "Arial", color: GRAY }),
          new ExternalHyperlink({
            link: "https://www.linkedin.com/in/matthew-carullo",
            children: [new TextRun({ text: "linkedin.com/in/matthew-carullo", style: "Hyperlink", size: 18, font: "Arial" })]
          }),
          new TextRun({ text: "  |  ", size: 18, font: "Arial", color: GRAY }),
          new ExternalHyperlink({
            link: "https://github.com/mcarullo-tech",
            children: [new TextRun({ text: "github.com/mcarullo-tech", style: "Hyperlink", size: 18, font: "Arial" })]
          })
        ]
      }),

      // SUMMARY
      sectionHeader("Summary"),
      new Paragraph({
        spacing: { before: 55, after: 55 },
        children: [
          new TextRun({
            text: "Mechanical & automation engineer with hands-on experience across manufacturing, mechanical design, and automation in industrial environments. Combines manufacturing experience from Tesla with Python-based signal processing and robotics automation work at Hatch. Comfortable operating at the intersection of mechanical systems, data pipelines, and process improvement.",
            size: 21, font: "Arial", color: BLACK
          })
        ]
      }),

      // EXPERIENCE
      sectionHeader("Experience"),

      // HATCH
      jobHeader("Mechanical & Automation Engineer", "Hatch Ltd.", "Mississauga, ON", "Aug 2024 – Present"),
      bullet("Built an automated Python signal-processing pipeline to replace a fully manual FFT interpretation workflow, implementing dominant frequency detection and classification, error analysis, and automated reporting, cutting analysis time by 40% and improving defect classification accuracy by 30%."),
      bullet("Developed an end-to-end business case for robotic automation deployment in harsh-environment NDT, encompassing technical requirements, financial model, and executive pitch deck, targeting 50% reduction in manual intervention."),
      bullet("Designed and prototyped automation solutions for harsh-environment inspection campaigns, working across global client sites under tight operational deadlines."),

      // TESLA MARKHAM
      jobHeader("Mechanical Engineer", "Tesla Inc.", "Markham, ON", "May 2023 – Aug 2023"),
      bullet("Engineered mechanical design revisions for high-speed lithium-ion cell manufacturing lines using SolidWorks and GD&T, standardizing improvements across global automation platforms via PLM systems."),
      bullet("Diagnosed excessive backlash and vibration in a high-throughput infeed gearbox through dynamic analysis and root-cause investigation, driving a mechanical redesign that improved powertrain precision and cut cycle-time variability by 20%."),
      bullet("Designed a custom in-line inspection tool for mechanical sealing machinery, enabling real-time seal validation and supporting a 25% boost in QA throughput."),

      // TESLA PALO ALTO
      jobHeader("Mechanical Engineer", "Tesla Inc.", "Palo Alto, CA", "Sept 2022 – Dec 2022"),
      bullet("Performed FMEA on Semi-truck cabin components, qualifying water ingress risk and informing design changes ahead of launch."),
      bullet("Designed a 3D-printable drip guard with full DFM analysis for scalable production, preventing an estimated $600K in electronics damage across early fleet builds."),

      // TESLA FREMONT
      jobHeader("Manufacturing Engineer", "Tesla Inc.", "Fremont, CA", "Jan 2022 – Apr 2022"),
      bullet("Led process improvement studies on laser welding systems in lithium-ion cell assembly, increasing machine yield from 92% to 99.8% through targeted validation and equipment optimization."),
      bullet("Prototyped a machine-learning vision system for automated weld classification, applying feature extraction and supervised learning techniques to reduce false negatives."),
      bullet("Implemented offline quality control measures across key assembly line stages, improving defect detection and production conformance."),

      // SKILLS
      sectionHeader("Technical Skills"),
      skillRow("Software & Data", "Python (SciPy, NumPy, Pandas), Algorithm Development, Data Visualization, SQL, MATLAB"),
      skillRow("Mechanical Design", "SolidWorks, AutoCAD, GD&T, Tolerance Stackup Analysis, DFM/DFA, PLM Systems"),
      skillRow("Analysis & Simulation", "ANSYS (FEA), Design of Experiments, Statistical Process Control"),
      skillRow("Manufacturing", "Laser Welding, CNC Machining, Injection Molding, 3D Printing, Rapid Prototyping"),
      skillRow("Electronics & Embedded", "Microcontrollers (Arduino), Sensor Integration, Actuator & Motor Control"),

      // EDUCATION
      sectionHeader("Education"),
      jobHeader("BASc, Mechanical Engineering", "University of Waterloo", "Waterloo, ON", "May 2024"),
      bullet("4.0 GPA, Dean's Honours List"),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/mnt/user-data/outputs/Matthew_Carullo_Resume.docx", buffer);
  console.log("Done");
});
