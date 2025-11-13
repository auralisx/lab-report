/** Lab Test Objects with placeholder and result */

// Update Patient Info in the Header
function replacePlaceholdersInHeader(doc, formData) {
  const header = doc.getHeader();
  if (!header) {
    Logger.log("No header found in the document.");
    return;
  }

  const placeholders = {
    "{{Patient_Name}}": formData[0],
    "{{Age}}": formData[5],
    "{{Gender}}": formData[6],
    "{{Address}}": formData[1],
    "{{Phone_No}}": formData[2],
    "{{Lab_No}}": formData[89],
    "{{Collection_Date}}": [formData[7], formData[8]].join(" ").toUpperCase(),
    "{{Dispatch_Date}}": formattedDate,
    "{{Referral}}": formData[3],
    "{{Patient_No}}": formData[4],
  };

  Object.entries(placeholders).forEach(([placeholder, value]) => {
    header.replaceText(placeholder, value || "");
  });
}

function getLabTests(formData) {
  return {
    Hematology: {
      department: "HEMATOLOGY",
      tableHeading: "HEMATOLOGY - CBC",
      templateTableIndex: 0,
      tests: [
        { placeholder: "<<TLC>>", result: formData[11] },
        { placeholder: "<<Neutrophil>>", result: formData[12] },
        { placeholder: "<<Lymphocyte>>", result: formData[13] },
        { placeholder: "<<Monocyte>>", result: formData[14] },
        { placeholder: "<<Eosinophil>>", result: formData[15] },
        { placeholder: "<<Basophil>>", result: formData[16] },
        { placeholder: "<<Hemoglobin>>", result: formData[17] },
        { placeholder: "<<Platelets>>", result: formData[18] },
        { placeholder: "<<RBC>>", result: formData[19] },
        { placeholder: "<<PCV>>", result: formData[20] },
        { placeholder: "<<MCV>>", result: formData[21] },
        { placeholder: "<<MCH>>", result: formData[22] },
        { placeholder: "<<MCHC>>", result: formData[23] },
      ],
    },
    Lipid_Profile: {
      department: "BIO-CHEMISTRY",
      tableHeading: "BIO-CHEMISTRY - LIPID PROFILE",
      templateTableIndex: 1,
      tests: [
        { placeholder: "<<Triglyceride>>", result: formData[24] },
        { placeholder: "<<Cholesterol>>", result: formData[25] },
        { placeholder: "<<HDL Cholesterol>>", result: formData[26] },
        { placeholder: "<<LDL Cholesterol>>", result: formData[27] },
        { placeholder: "<<VLDL>>", result: formData[28] },
        { placeholder: "<<Non-HDL>>", result: formData[78] },
      ],
    },
    LFT: {
      department: "BIO-CHEMISTRY",
      tableHeading: "BIO-CHEMISTRY - LIVER FUNCTION TEST",
      templateTableIndex: 2,
      tests: [
        { placeholder: "<<Total Bilirubin>>", result: formData[29] },
        { placeholder: "<<Direct Bilirubin>>", result: formData[30] },
        { placeholder: "<<Alkaline Phosphatase>>", result: formData[31] },
        { placeholder: "<<SGPT/ALT>>", result: formData[32] },
        { placeholder: "<<SGOT/AST>>", result: formData[33] },
        { placeholder: "<<Total Protein>>", result: formData[34] },
        { placeholder: "<<Albumin>>", result: formData[35] },
        { placeholder: "<<AG Ratio>>", result: formData[36] },
        { placeholder: "<<Globulin>>", result: formData[79] },
      ],
    },
    RFT: {
      department: "BIO-CHEMISTRY",
      tableHeading: "BIO-CHEMISTRY - RENAL FUNCTION TEST",
      templateTableIndex: 3,
      tests: [
        { placeholder: "<<Urea>>", result: formData[40] },
        { placeholder: "<<Creatinine>>", result: formData[41] },
        { placeholder: "<<Sodium>>", result: formData[42] },
        { placeholder: "<<Potassium>>", result: formData[43] },
      ],
    },
    Stool_RE: {
      department: "PARASITOLOGY",
      tableHeading: "PARASITOLOGY - STOOL ROUTINE EXAMINATION",
      templateTableIndex: 4,
      tests: [
        { placeholder: "<<Colour>>", result: formData[44] },
        { placeholder: "<<Consistency>>", result: formData[45] },
        { placeholder: "<<Mucus>>", result: formData[46] },
        { placeholder: "<<Blood>>", result: formData[47] },
        // { placeholder: '<<obt>>', result: formData[98]},
        { placeholder: "<<Cyst of Parasite>>", result: formData[48] },
        { placeholder: "<<Ova of Parasite>>", result: formData[49] },
        { placeholder: "<<Pus cells>>", result: formData[50] },
        { placeholder: "<<Red cells>>", result: formData[51] },
        { placeholder: "<<Yeast Cells>>", result: formData[52] },
        { placeholder: "<<Others>>", result: formData[53] },
        { placeholder: "<<ufp>>", result: formData[88] },
        { placeholder: "<<cyst>>", result: formData[94] },
      ],
    },
    Urine_RE: {
      department: "PARASITOLOGY",
      tableHeading: "PARASITOLOGY - URINE ROUTINE EXAMINATION",
      templateTableIndex: 5,
      tests: [
        { placeholder: "<<Colour>>", result: formData[54] },
        { placeholder: "<<Appearance>>", result: formData[55] },
        { placeholder: "<<PH>>", result: formData[56] },
        { placeholder: "<<Sugar>>", result: formData[57] },
        { placeholder: "<<Albumin>>", result: formData[58] },
        { placeholder: "<<Pus cells>>", result: formData[59] },
        { placeholder: "<<Red cells>>", result: formData[60] },
        { placeholder: "<<Epithelial cells>>", result: formData[61] },
        { placeholder: "<<Yeast Cells>>", result: formData[62] },
        { placeholder: "<<Calcium Oxalate Crystals>>", result: formData[63] },
        { placeholder: "<<Amorphous Urates>>", result: formData[64] },
        { placeholder: "<<Amorphous Phosphates>>", result: formData[65] },
        { placeholder: "<<Granular Casts>>", result: formData[66] },
        { placeholder: "<<Others>>", result: formData[67] },
      ],
    },
    Urine_Culture: {
      department: "MICROBIOLOGY",
      tableHeading: "MICROBIOLOGY - URINE CULTURE",
      templateTableIndex: 6,
      tests: [{ placeholder: "<<Result>>", result: formData[75] }],
    },
    Thyroid_Function: {
      department: "IMMUNOLOGY",
      tableHeading: "IMMUNOLOGY - THYROID FUNCTION TEST",
      templateTableIndex: 7,
      commentTableIndex: 1,
      tests: [
        { placeholder: "<<FT3>>", result: formData[37] },
        { placeholder: "<<FT4>>", result: formData[38] },
        { placeholder: "<<TSH>>", result: formData[39] },
      ],
    },
    Iron_Profile: {
      department: "BIO-CHEMISTRY",
      tableHeading: "BIO-CHEMISTRY - IRON PROFILE",
      templateTableIndex: 8,
      tests: [
        { placeholder: "<<Iron>>", result: formData[71] },
        { placeholder: "<<TIBC>>", result: formData[72] },
        { placeholder: "<<UIBC>>", result: formData[73] },
        { placeholder: "<<TS>>", result: formData[74] },
      ],
    },
    Dengue: {
      department: "SEROLOGY",
      tableHeading: "DENGUE COMBO SEROLOGY SCREEN",
      templateTableIndex: 9,
      tests: [
        { placeholder: "<<IgG>>", result: formData[68] },
        { placeholder: "<<IgM>>", result: formData[69] },
        { placeholder: "<<NS1>>", result: formData[70] },
      ],
    },
    URIC_ACID: {
      department: "BIO-CHEMISTRY",
      tableHeading: "BIO-CHEMISTRY - URIC ACID",
      templateTableIndex: 10,
      tests: [{ placeholder: "<<Uric Acid>>", result: formData[87] }],
    },
    Vitamin_D3: {
      department: "IMMUNOLOGY",
      tableHeading: "IMMUNOLOGY - VITAMIN D3",
      templateTableIndex: 11,
      commentTableIndex: 4,
      tests: [{ placeholder: "<<VD>>", result: formData[85] }],
    },
    Vitamin_B12: {
      department: "IMMUNOLOGY",
      tableHeading: "IMMUNOLOGY - VITAMIN B12",
      templateTableIndex: 12,
      commentTableIndex: 5,
      tests: [{ placeholder: "<<VB>>", result: formData[84] }],
    },
    Blood_Sugar: {
      department: "BIO-CHEMISTRY",
      tableHeading: "BIO-CHEMISTRY - BLOOD SUGAR",
      templateTableIndex: 13,
      tests: [
        { placeholder: "<<random>>", result: formData[81] },
        { placeholder: "<<fasting>>", result: formData[82] },
        { placeholder: "<<pp>>", result: formData[83] },
      ],
    },
    Calcium: {
      department: "BIO-CHEMISTRY",
      tableHeading: "BIO-CHEMISTRY - CALCIUM",
      templateTableIndex: 14,
      tests: [{ placeholder: "<<Calcium>>", result: formData[80] }],
    },
    HbA1c: {
      department: "BIO-CHEMISTRY",
      tableHeading: "BIO-CHEMISTRY - HbA1c",
      templateTableIndex: 15,
      commentTableIndex: 7,
      tests: [{ placeholder: "<<HbA1c>>", result: formData[86] }],
    },
    TSH: {
      department: "IMMUNOLOGY",
      tableHeading: "IMMUNOLOGY - TSH",
      templateTableIndex: 16,
      commentTableIndex: 6,
      tests: [{ placeholder: "<<TSH>>", result: formData[90] }],
    },
    SPOT: {
      department: "BIO-CHEMISTRY",
      tableHeading: "SPOT A:C",
      templateTableIndex: 17,
      tests: [
        { placeholder: "<<A/M>>", result: formData[91] },
        { placeholder: "<<Urine>>", result: formData[92] },
        { placeholder: "<<ACR>>", result: formData[93] },
      ],
    },
    OBT: {
      department: "MICROBIOLOGY",
      tableHeading: "MICROBIOLOGY - STOOL OCCULT BLOOD",
      templateTableIndex: 18,
      tests: [{ placeholder: "<<OBT>>", result: formData[94] }],
    },
    IPTH: {
      department: "IMMUNOLOGY",
      tableHeading: "INTACT PARATHYROID HORMONE",
      templateTableIndex: 19,
      commentTableIndex: 8,
      tests: [{ placeholder: "<<IPTH>>", result: formData[95] }],
    },
    Amylase: {
      department: "BIO-CHEMISTRY",
      tableHeading: "BIO-CHEMISTRY - SERUM AMYLASE",
      templateTableIndex: 20,
      tests: [{ placeholder: "<<Amylase>>", result: formData[97] }],
    },
    Serology: {
      department: "SEROLOGY",
      tableHeading: "SEROLOGY",
      templateTableIndex: 21,
      commentTableIndex: 9,
      tests: [
        { placeholder: "<<VDRL>>", result: formData[104] },
        { placeholder: "<<HIV>>", result: formData[103] },
        { placeholder: "<<HCV>>", result: formData[102] },
        { placeholder: "<<HBsAg>>", result: formData[101] },
        { placeholder: "<<TPHA>>", result: formData[100] },
      ],
    },
    VDRL: {
      department: "SEROLOGY",
      tableHeading: "SEROLOGY - VDRL",
      templateTableIndex: 23,
      commentTableIndex: 9,
      tests: [{ placeholder: "<<VDRL>>", result: formData[104] }],
    },
    HIV: {
      department: "SEROLOGY",
      tableHeading: "SEROLOGY - HIV",
      templateTableIndex: 24,
      tests: [{ placeholder: "<<HIV>>", result: formData[103] }],
    },
    HCV: {
      department: "SEROLOGY",
      tableHeading: "SEROLOGY - HCV",
      templateTableIndex: 25,
      tests: [{ placeholder: "<<HCV>>", result: formData[102] }],
    },
    HBsAg: {
      department: "SEROLOGY",
      tableHeading: "SEROLOGY - HBsAg",
      templateTableIndex: 26,
      tests: [{ placeholder: "<<HBsAg>>", result: formData[101] }],
    },
    TPHA: {
      department: "SEROLOGY",
      tableHeading: "SEROLOGY - TPHA",
      templateTableIndex: 27,
      tests: [{ placeholder: "<<TPHA>>", result: formData[100] }],
    },
    UPT: {
      department: "SEROLOGY",
      tableHeading: "SEROLOGY - UPT",
      templateTableIndex: 22,
      tests: [{ placeholder: "<<UPT>>", result: formData[99] }],
    },
    // ... other test types ...
  };
}
