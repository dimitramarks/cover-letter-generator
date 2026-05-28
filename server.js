require('dotenv').config({ path: require('path').join(__dirname, '.env') });
const express   = require('express');
const multer    = require('multer');
const path      = require('path');
const pdfParse  = require('pdf-parse');
const mammoth   = require('mammoth');
const {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, ImageRun, WidthType, AlignmentType, BorderStyle,
  HeadingLevel, UnderlineType, VerticalAlign, TabStopType, TabStopPosition,
} = require('docx');
const { createClient } = require('@supabase/supabase-js');

const fs     = require('fs');
const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
app.use(express.json({ limit: '5mb' }));
app.use(express.static(__dirname));

const PORT = process.env.TRP_PORT || 8081;
const API_KEY = process.env.ANTHROPIC_API_KEY;

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

// ─────────────────────────────────────────────
//  STATUTE HISTORY
//  Each offence maps to an array of time-versioned
//  statute entries. fromDate is inclusive,
//  toDate is exclusive (null = still current).
//  Per Canada (Citizenship and Immigration) v. Tran,
//  2010 SCC 58, equivalency is assessed as at the
//  date of commission of the foreign offence.
// ─────────────────────────────────────────────
const STATUTE_HISTORY = {

  // ── DRIVING OFFENCES — restructured by Bill C-46 on December 18, 2018 ──
  // Pre-2018: hybrid, max 5 years indictable → s.36(2) IRPA (criminality)
  // Post-2018: hybrid, max 10 years indictable → s.36(1) IRPA (serious criminality)

  'DUI / DWI / Impaired Driving (Alcohol)': [
    {
      fromDate: '1900-01-01', toDate: '2018-12-18',
      statute: 'Section 253 of the Criminal Code of Canada',
      title: 'Operation While Impaired',
      hybrid: true, maxSummary: '18 months imprisonment', maxIndictable: '5 years imprisonment',
      irpa: 's.36(2) — criminality (not serious criminality)',
      irpaThresholdCrossed: true,
    },
    {
      fromDate: '2018-12-18', toDate: null,
      statute: 'Section 320.14 of the Criminal Code of Canada',
      title: 'Operation While Impaired',
      hybrid: true, maxSummary: '2 years less a day imprisonment', maxIndictable: '10 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'DUI / DWI / Impaired Driving (Drugs)': [
    {
      fromDate: '1900-01-01', toDate: '2018-12-18',
      statute: 'Section 253 of the Criminal Code of Canada',
      title: 'Operation While Impaired',
      hybrid: true, maxSummary: '18 months imprisonment', maxIndictable: '5 years imprisonment',
      irpa: 's.36(2) — criminality (not serious criminality)',
      irpaThresholdCrossed: true,
    },
    {
      fromDate: '2018-12-18', toDate: null,
      statute: 'Section 320.14 of the Criminal Code of Canada',
      title: 'Operation While Impaired',
      hybrid: true, maxSummary: '2 years less a day imprisonment', maxIndictable: '10 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'DUI / DWI / Impaired Driving — Causing Bodily Harm': [
    {
      fromDate: '1900-01-01', toDate: '2018-12-18',
      statute: 'Section 255(2) of the Criminal Code of Canada',
      title: 'Impaired Driving Causing Bodily Harm',
      indictable: true, maxIndictable: '10 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
    {
      fromDate: '2018-12-18', toDate: null,
      statute: 'Section 320.14(2) of the Criminal Code of Canada',
      title: 'Impaired Operation Causing Bodily Harm',
      hybrid: true, maxSummary: '2 years less a day imprisonment', maxIndictable: '14 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'DUI / DWI / Impaired Driving — Causing Death': [
    {
      fromDate: '1900-01-01', toDate: '2018-12-18',
      statute: 'Section 255(3) of the Criminal Code of Canada',
      title: 'Impaired Driving Causing Death',
      indictable: true, maxIndictable: 'Life imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
    {
      fromDate: '2018-12-18', toDate: null,
      statute: 'Section 320.14(3) of the Criminal Code of Canada',
      title: 'Impaired Operation Causing Death',
      indictable: true, maxIndictable: 'Life imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'Dangerous Driving / Dangerous Operation': [
    {
      fromDate: '1900-01-01', toDate: '2018-12-18',
      statute: 'Section 249 of the Criminal Code of Canada',
      title: 'Dangerous Operation of Motor Vehicles, Vessels and Aircraft',
      hybrid: true, maxSummary: '2 years imprisonment', maxIndictable: '5 years imprisonment',
      irpa: 's.36(2) — criminality (not serious criminality)',
      irpaThresholdCrossed: true,
    },
    {
      fromDate: '2018-12-18', toDate: null,
      statute: 'Section 320.13 of the Criminal Code of Canada',
      title: 'Dangerous Operation',
      hybrid: true, maxSummary: '2 years less a day imprisonment', maxIndictable: '10 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'Dangerous Operation — Causing Bodily Harm': [
    {
      fromDate: '1900-01-01', toDate: '2018-12-18',
      statute: 'Section 249(3) of the Criminal Code of Canada',
      title: 'Dangerous Operation Causing Bodily Harm',
      indictable: true, maxIndictable: '10 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
    {
      fromDate: '2018-12-18', toDate: null,
      statute: 'Section 320.13(2) of the Criminal Code of Canada',
      title: 'Dangerous Operation Causing Bodily Harm',
      hybrid: true, maxSummary: '2 years less a day imprisonment', maxIndictable: '14 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'Dangerous Operation — Causing Death': [
    {
      fromDate: '1900-01-01', toDate: '2018-12-18',
      statute: 'Section 249(4) of the Criminal Code of Canada',
      title: 'Dangerous Operation Causing Death',
      indictable: true, maxIndictable: '14 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
    {
      fromDate: '2018-12-18', toDate: null,
      statute: 'Section 320.13(3) of the Criminal Code of Canada',
      title: 'Dangerous Operation Causing Death',
      indictable: true, maxIndictable: 'Life imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'Failure to Stop / Hit and Run': [
    {
      fromDate: '1900-01-01', toDate: '2018-12-18',
      statute: 'Section 252 of the Criminal Code of Canada',
      title: 'Failure to Stop at Scene of Accident',
      hybrid: true, maxSummary: '2 years imprisonment', maxIndictable: '5 years imprisonment',
      irpa: 's.36(2) — criminality (not serious criminality)',
      irpaThresholdCrossed: true,
    },
    {
      fromDate: '2018-12-18', toDate: null,
      statute: 'Section 320.16 of the Criminal Code of Canada',
      title: 'Failure to Stop after Accident',
      hybrid: true, maxSummary: '2 years less a day imprisonment', maxIndictable: '10 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'Failure to Stop / Hit and Run — Causing Bodily Harm': [
    {
      fromDate: '1900-01-01', toDate: '2018-12-18',
      statute: 'Section 252(1.2) of the Criminal Code of Canada',
      title: 'Failure to Stop at Scene of Accident Causing Bodily Harm',
      indictable: true, maxIndictable: '10 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
    {
      fromDate: '2018-12-18', toDate: null,
      statute: 'Section 320.16(2) of the Criminal Code of Canada',
      title: 'Failure to Stop after Accident Causing Bodily Harm',
      hybrid: true, maxSummary: '2 years less a day imprisonment', maxIndictable: '14 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'Failure to Stop / Hit and Run — Causing Death': [
    {
      fromDate: '1900-01-01', toDate: '2018-12-18',
      statute: 'Section 252(1.3) of the Criminal Code of Canada',
      title: 'Failure to Stop at Scene of Accident Causing Death',
      indictable: true, maxIndictable: 'Life imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
    {
      fromDate: '2018-12-18', toDate: null,
      statute: 'Section 320.16(3) of the Criminal Code of Canada',
      title: 'Failure to Stop after Accident Causing Death',
      indictable: true, maxIndictable: 'Life imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'Driving While Prohibited / Suspended Licence': [
    {
      fromDate: '1900-01-01', toDate: '2018-12-18',
      statute: 'Section 259(4) of the Criminal Code of Canada',
      title: 'Driving While Disqualified',
      hybrid: true, maxSummary: '2 years imprisonment', maxIndictable: '5 years imprisonment',
      irpa: 's.36(2) — criminality (not serious criminality)',
      irpaThresholdCrossed: true,
    },
    {
      fromDate: '2018-12-18', toDate: null,
      statute: 'Section 320.18(1) of the Criminal Code of Canada',
      title: 'Operation While Prohibited',
      hybrid: true, maxSummary: '2 years less a day imprisonment', maxIndictable: '10 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'Operation While Prohibited': [{ fromDate: '2018-12-18', toDate: null,
    statute: 'Section 320.18(1) of the Criminal Code of Canada',
    title: 'Operation While Prohibited',
    hybrid: true, maxSummary: '2 years less a day imprisonment', maxIndictable: '10 years imprisonment',
    irpa: 's.36(1) — serious criminality',
  }],

  // ── ALL OTHER OFFENCES — max punishments stable since 2002 ──

  'Simple Assault / Battery': [{ fromDate: '1900-01-01', toDate: null,
    statute: 'Section 266 of the Criminal Code of Canada', title: 'Assault',
    hybrid: true, maxSummary: '18 months imprisonment', maxIndictable: '5 years imprisonment',
    irpa: 's.36(2) — criminality',
  }],

  'Assault Causing Bodily Harm': [{ fromDate: '1900-01-01', toDate: null,
    statute: 'Section 267 of the Criminal Code of Canada', title: 'Assault with a Weapon or Causing Bodily Harm',
    hybrid: true, maxSummary: '18 months imprisonment', maxIndictable: '10 years imprisonment',
    irpa: 's.36(1) — serious criminality',
  }],

  'Aggravated Assault': [{ fromDate: '1900-01-01', toDate: null,
    statute: 'Section 268 of the Criminal Code of Canada', title: 'Aggravated Assault',
    indictable: true, maxIndictable: '14 years imprisonment',
    irpa: 's.36(1) — serious criminality',
  }],

  'Sexual Assault': [{ fromDate: '1900-01-01', toDate: null,
    statute: 'Section 271 of the Criminal Code of Canada', title: 'Sexual Assault',
    hybrid: true, maxSummary: '18 months imprisonment', maxIndictable: '10 years imprisonment',
    irpa: 's.36(1) — serious criminality',
  }],

  'Theft Under $5,000': [{ fromDate: '1900-01-01', toDate: null,
    statute: 'Section 334(b) of the Criminal Code of Canada', title: 'Theft',
    hybrid: true, maxSummary: '2 years imprisonment', maxIndictable: '2 years imprisonment',
    irpa: 's.36(2) — criminality',
  }],

  'Theft Over $5,000': [{ fromDate: '1900-01-01', toDate: null,
    statute: 'Section 334(a) of the Criminal Code of Canada', title: 'Theft',
    indictable: true, maxIndictable: '10 years imprisonment',
    irpa: 's.36(1) — serious criminality',
  }],

  'Fraud Under $5,000': [{ fromDate: '1900-01-01', toDate: null,
    statute: 'Section 380(1)(b) of the Criminal Code of Canada', title: 'Fraud',
    hybrid: true, maxSummary: '2 years imprisonment', maxIndictable: '2 years imprisonment',
    irpa: 's.36(2) — criminality',
  }],

  'Fraud Over $5,000': [
    {
      fromDate: '1900-01-01', toDate: '2004-09-15',
      statute: 'Section 380(1)(a) of the Criminal Code of Canada', title: 'Fraud',
      indictable: true, maxIndictable: '10 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
    {
      fromDate: '2004-09-15', toDate: null,
      statute: 'Section 380(1)(a) of the Criminal Code of Canada', title: 'Fraud',
      indictable: true, maxIndictable: '14 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  // ── CANNABIS / MARIJUANA OFFENCES ──────────────────────────────────────────
  // Pre-Sep 3, 1997:   Narcotic Control Act (R.S.C. 1985, c. N-1) — cannabis was in NCA Schedule item 3
  // Sep 3, 1997 – Oct 17, 2018: CDSA Schedule II (S.C. 1996, c. 19)
  // Post-Oct 17, 2018: Cannabis Act (S.C. 2018, c. 16)
  // Cannabis production also has a mid-point: Bill C-10 (S.C. 2012, c. 1, in force Nov 6, 2012)
  //   raised CDSA Schedule II production max from 7 years → 14 years, crossing the IRPA s.36(1) threshold.

  'Marijuana / Cannabis — Simple Possession': [
    { fromDate: '1961-01-01', toDate: '1997-09-03',
      statute: 'Section 3(1) of the Narcotic Control Act (R.S.C. 1985, c. N-1)', title: 'Possession of Narcotic (Cannabis)',
      hybrid: true, maxSummary: '6 months imprisonment (or $1,000 fine)', maxIndictable: '7 years imprisonment',
      irpa: 's.36(2) — criminality',
      amendmentLabel: 'Controlled Drugs and Substances Act (S.C. 1996, c. 19)',
    },
    { fromDate: '1997-09-03', toDate: '2018-10-17',
      statute: 'Section 4(1) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Possession of Schedule II Substance (Cannabis)',
      hybrid: true, maxSummary: '6 months imprisonment', maxIndictable: '5 years less a day imprisonment',
      irpa: 's.36(2) — criminality',
      amendmentLabel: 'Cannabis Act (S.C. 2018, c. 16)',
    },
    { fromDate: '2018-10-17', toDate: null,
      statute: 'Section 8 of the Cannabis Act (S.C. 2018, c. 16)', title: 'Possession of Cannabis',
      hybrid: true, maxSummary: '6 months imprisonment (or $5,000 fine)', maxIndictable: '5 years less a day imprisonment',
      irpa: 's.36(2) — criminality',
    },
  ],

  'Marijuana / Cannabis — Trafficking or Distribution': [
    { fromDate: '1961-01-01', toDate: '1997-09-03',
      statute: 'Section 4 of the Narcotic Control Act (R.S.C. 1985, c. N-1)', title: 'Trafficking in Narcotic (Cannabis)',
      indictable: true, maxIndictable: 'Life imprisonment',
      irpa: 's.36(1) — serious criminality',
      amendmentLabel: 'Controlled Drugs and Substances Act (S.C. 1996, c. 19)',
    },
    { fromDate: '1997-09-03', toDate: '2018-10-17',
      statute: 'Section 5(1) and/or 5(2) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Trafficking in Schedule II Substance (Cannabis)',
      indictable: true, maxIndictable: 'Life imprisonment',
      irpa: 's.36(1) — serious criminality',
      amendmentLabel: 'Cannabis Act (S.C. 2018, c. 16)',
    },
    { fromDate: '2018-10-17', toDate: null,
      statute: 'Section 9 and/or Section 10 of the Cannabis Act (S.C. 2018, c. 16)', title: 'Distribution or Selling of Cannabis',
      hybrid: true, maxSummary: '6 months imprisonment (or $5,000 fine)', maxIndictable: '14 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'Marijuana / Cannabis — Import or Export': [
    { fromDate: '1961-01-01', toDate: '1997-09-03',
      statute: 'Section 5 of the Narcotic Control Act (R.S.C. 1985, c. N-1)', title: 'Importing or Exporting Narcotic (Cannabis)',
      indictable: true, maxIndictable: 'Life imprisonment (minimum 7 years)',
      irpa: 's.36(1) — serious criminality',
      amendmentLabel: 'Controlled Drugs and Substances Act (S.C. 1996, c. 19)',
    },
    { fromDate: '1997-09-03', toDate: '2018-10-17',
      statute: 'Section 6(1) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Importing and Exporting Schedule II Substance (Cannabis)',
      indictable: true, maxIndictable: 'Life imprisonment',
      irpa: 's.36(1) — serious criminality',
      amendmentLabel: 'Cannabis Act (S.C. 2018, c. 16)',
    },
    { fromDate: '2018-10-17', toDate: null,
      statute: 'Section 11 of the Cannabis Act (S.C. 2018, c. 16)', title: 'Importing and Exporting Cannabis',
      hybrid: true, maxSummary: '6 months imprisonment (or $5,000 fine)', maxIndictable: '14 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  // Four versions: NCA cultivation (max 7y), CDSA pre-2012 (max 7y), CDSA post-2012 (max 14y, threshold crossed), Cannabis Act.
  'Marijuana / Cannabis — Production or Cultivation': [
    { fromDate: '1961-01-01', toDate: '1997-09-03',
      statute: 'Section 6 of the Narcotic Control Act (R.S.C. 1985, c. N-1)', title: 'Cultivation of Marihuana or Opium Poppy',
      indictable: true, maxIndictable: '7 years imprisonment',
      irpa: 's.36(2) — criminality',
      irpaThresholdCrossed: true,
      amendmentLabel: 'Controlled Drugs and Substances Act (S.C. 1996, c. 19)',
    },
    { fromDate: '1997-09-03', toDate: '2012-11-06',
      statute: 'Section 7(1) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Production of Schedule II Substance (Cannabis)',
      indictable: true, maxIndictable: '7 years imprisonment',
      irpa: 's.36(2) — criminality',
      irpaThresholdCrossed: true,
      amendmentLabel: 'Bill C-10 (Safe Streets and Communities Act, S.C. 2012, c. 1)',
    },
    { fromDate: '2012-11-06', toDate: '2018-10-17',
      statute: 'Section 7(1) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Production of Schedule II Substance (Cannabis)',
      indictable: true, maxIndictable: '14 years imprisonment',
      irpa: 's.36(1) — serious criminality',
      amendmentLabel: 'Cannabis Act (S.C. 2018, c. 16)',
    },
    { fromDate: '2018-10-17', toDate: null,
      statute: 'Section 12 of the Cannabis Act (S.C. 2018, c. 16)', title: 'Production of Cannabis',
      hybrid: true, maxSummary: '6 months imprisonment (or $100,000 fine)', maxIndictable: '14 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  // ── NON-CANNABIS DRUG OFFENCES ─────────────────────────────────────────────
  // Narcotics (heroin, cocaine, opioids, PCP, fentanyl) were governed by the Narcotic Control Act
  // until Sep 3, 1997, then by the CDSA Schedule I.
  // Amphetamines (meth, MDMA) and hallucinogens (LSD, psilocybin) were governed by
  // the Food and Drug Act (Schedules G and H) pre-1997, then CDSA Schedules I and III.
  // Benzodiazepines/barbiturates/steroids (Schedule IV) have no possession offence under CDSA.

  // ─ POSSESSION — Schedules I, II, III only (no Schedule IV possession offence) ─

  'Possession — Schedule I Controlled Substance (Cocaine / Heroin / Fentanyl / Opioids / PCP)': [
    { fromDate: '1961-01-01', toDate: '1997-09-03',
      statute: 'Section 3(1) of the Narcotic Control Act (R.S.C. 1985, c. N-1)', title: 'Possession of Narcotic',
      hybrid: true, maxSummary: '6 months imprisonment (or $1,000 fine)', maxIndictable: '7 years imprisonment',
      irpa: 's.36(2) — criminality',
    },
    { fromDate: '1997-09-03', toDate: null,
      statute: 'Section 4(1) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Possession of Schedule I Substance',
      hybrid: true, maxSummary: '6 months imprisonment (or $1,000 fine)', maxIndictable: '7 years imprisonment',
      irpa: 's.36(2) — criminality',
    },
  ],

  'Possession — Schedule II Controlled Substance (Synthetic Cannabinoids / K2 / Spice)': [{ fromDate: '1997-09-03', toDate: null,
    statute: 'Section 4(1) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Possession of Schedule II Substance',
    hybrid: true, maxSummary: '6 months imprisonment (or $1,000 fine)', maxIndictable: '5 years less a day imprisonment',
    irpa: 's.36(2) — criminality',
  }],

  'Possession — Schedule III Controlled Substance (LSD / Psilocybin / Mescaline / Methylphenidate)': [{ fromDate: '1997-09-03', toDate: null,
    statute: 'Section 4(1) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Possession of Schedule III Substance',
    hybrid: true, maxSummary: '6 months imprisonment (or $1,000 fine)', maxIndictable: '3 years imprisonment',
    irpa: 's.36(2) — criminality',
  }],

  // ─ TRAFFICKING — all four Schedules ─

  'Trafficking — Schedule I Controlled Substance (Cocaine / Heroin / Fentanyl / Opioids / Methamphetamine / MDMA)': [
    { fromDate: '1961-01-01', toDate: '1997-09-03',
      statute: 'Section 4 of the Narcotic Control Act (R.S.C. 1985, c. N-1)', title: 'Trafficking in Narcotic',
      indictable: true, maxIndictable: 'Life imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
    { fromDate: '1997-09-03', toDate: null,
      statute: 'Section 5(1) and/or 5(2) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Trafficking in Schedule I Substance',
      indictable: true, maxIndictable: 'Life imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'Trafficking — Schedule II Controlled Substance (Synthetic Cannabinoids / K2 / Spice)': [{ fromDate: '1997-09-03', toDate: null,
    statute: 'Section 5(1) and/or 5(2) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Trafficking in Schedule II Substance',
    indictable: true, maxIndictable: 'Life imprisonment',
    irpa: 's.36(1) — serious criminality',
  }],

  'Trafficking — Schedule III Controlled Substance (LSD / Psilocybin / Mescaline)': [{ fromDate: '1997-09-03', toDate: null,
    statute: 'Section 5(1) and/or 5(2) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Trafficking in Schedule III Substance',
    hybrid: true, maxSummary: '18 months imprisonment', maxIndictable: '10 years imprisonment',
    irpa: 's.36(1) — serious criminality',
  }],

  'Trafficking — Schedule IV Controlled Substance (Benzodiazepines / Barbiturates / Anabolic Steroids)': [{ fromDate: '1997-09-03', toDate: null,
    statute: 'Section 5(1) and/or 5(2) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Trafficking in Schedule IV Substance',
    hybrid: true, maxSummary: '1 year imprisonment', maxIndictable: '3 years imprisonment',
    irpa: 's.36(2) — criminality',
  }],

  // ─ PRODUCTION / MANUFACTURE ─

  'Production / Manufacture — Schedule I Controlled Substance (Cocaine / Heroin / Fentanyl / Methamphetamine)': [
    { fromDate: '1961-01-01', toDate: '1997-09-03',
      statute: 'Section 4 of the Narcotic Control Act (R.S.C. 1985, c. N-1)', title: 'Trafficking in Narcotic (Production / Manufacture)',
      indictable: true, maxIndictable: 'Life imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
    { fromDate: '1997-09-03', toDate: null,
      statute: 'Section 7(1) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Production of Schedule I Substance',
      indictable: true, maxIndictable: 'Life imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'Production / Manufacture — Schedule II Controlled Substance (Synthetic Cannabinoids)': [{ fromDate: '1997-09-03', toDate: null,
    statute: 'Section 7(1) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Production of Schedule II Substance',
    indictable: true, maxIndictable: 'Life imprisonment',
    irpa: 's.36(1) — serious criminality',
  }],

  'Production / Manufacture — Schedule III Controlled Substance (LSD / Psilocybin / Mescaline)': [{ fromDate: '1997-09-03', toDate: null,
    statute: 'Section 7(1) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Production of Schedule III Substance',
    hybrid: true, maxSummary: '18 months imprisonment', maxIndictable: '10 years imprisonment',
    irpa: 's.36(1) — serious criminality',
  }],

  'Production / Manufacture — Schedule IV Controlled Substance (Benzodiazepines / Barbiturates / Anabolic Steroids)': [{ fromDate: '1997-09-03', toDate: null,
    statute: 'Section 7(1) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Production of Schedule IV Substance',
    hybrid: true, maxSummary: '1 year imprisonment', maxIndictable: '3 years imprisonment',
    irpa: 's.36(2) — criminality',
  }],

  // ─ IMPORT / EXPORT ─

  'Import / Export — Schedule I or II Controlled Substance (Cocaine / Heroin / Fentanyl / Opioids)': [
    { fromDate: '1961-01-01', toDate: '1997-09-03',
      statute: 'Section 5 of the Narcotic Control Act (R.S.C. 1985, c. N-1)', title: 'Importing or Exporting Narcotic',
      indictable: true, maxIndictable: 'Life imprisonment (minimum 7 years)',
      irpa: 's.36(1) — serious criminality',
    },
    { fromDate: '1997-09-03', toDate: null,
      statute: 'Section 6(1) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Importing and Exporting Schedule I or II Substance',
      indictable: true, maxIndictable: 'Life imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'Import / Export — Schedule III Controlled Substance (LSD / Psilocybin / Mescaline)': [{ fromDate: '1997-09-03', toDate: null,
    statute: 'Section 6(1) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Importing and Exporting Schedule III Substance',
    hybrid: true, maxSummary: '18 months imprisonment', maxIndictable: '10 years imprisonment',
    irpa: 's.36(1) — serious criminality',
  }],

  'Import / Export — Schedule IV Controlled Substance (Benzodiazepines / Barbiturates / Anabolic Steroids)': [{ fromDate: '1997-09-03', toDate: null,
    statute: 'Section 6(1) of the Controlled Drugs and Substances Act (S.C. 1996, c. 19)', title: 'Importing and Exporting Schedule IV Substance',
    hybrid: true, maxSummary: '1 year imprisonment', maxIndictable: '3 years imprisonment',
    irpa: 's.36(2) — criminality',
  }],

  'Domestic Violence': [{ fromDate: '1900-01-01', toDate: null,
    statute: 'Section 266 of the Criminal Code of Canada', title: 'Assault',
    hybrid: true, maxSummary: '18 months imprisonment', maxIndictable: '5 years imprisonment',
    irpa: 's.36(2) — criminality',
  }],

  'Breaking and Entering / Burglary': [{ fromDate: '1900-01-01', toDate: null,
    statute: 'Section 348 of the Criminal Code of Canada', title: 'Breaking and Entering',
    hybrid: true, maxSummary: '2 years imprisonment', maxIndictable: 'Life imprisonment (dwelling) or 10 years (other)',
    irpa: 's.36(1) — serious criminality',
  }],

  'Weapons Possession (Prohibited/Restricted)': [{ fromDate: '1900-01-01', toDate: null,
    statute: 'Section 91 and/or Section 92 of the Criminal Code of Canada', title: 'Unauthorized Possession of Firearm',
    hybrid: true, maxSummary: '12 months imprisonment', maxIndictable: '5 years imprisonment',
    irpa: 's.36(2) — criminality',
  }],

  'Disorderly Conduct / Causing a Disturbance': [{ fromDate: '1900-01-01', toDate: null,
    statute: 'Section 175 of the Criminal Code of Canada', title: 'Causing Disturbance, Indecent Exhibition, Loitering, etc.',
    summary: true, maxSummary: '2 years imprisonment',
    irpa: 's.36(2) — criminality',
  }],

  'Resisting Arrest / Obstruction of Justice': [{ fromDate: '1900-01-01', toDate: null,
    statute: 'Section 129 of the Criminal Code of Canada', title: 'Offences Relating to Public or Peace Officer',
    summary: true, maxSummary: '2 years imprisonment',
    irpa: 's.36(2) — criminality',
  }],

  'Mischief / Vandalism': [{ fromDate: '1900-01-01', toDate: null,
    statute: 'Section 430 of the Criminal Code of Canada', title: 'Mischief',
    hybrid: true, maxSummary: '2 years imprisonment', maxIndictable: '10 years (over $5,000) or 2 years (under $5,000)',
    irpa: 's.36(1) or s.36(2) depending on amount',
  }],

  'Incest': [{ fromDate: '1900-01-01', toDate: null,
    statute: 'Section 155 of the Criminal Code of Canada', title: 'Incest',
    indictable: true, maxIndictable: '14 years imprisonment',
    irpa: 's.36(1) — serious criminality',
  }],

  // ── CSAM — s.163.1 amended July 17, 2015: possession 5y → 10y; making/distribution 10y → 14y ──
  'Child Pornography / CSAM — Making or Distribution': [
    { fromDate: '1900-01-01', toDate: '2015-07-17',
      statute: 'Section 163.1(2) and (3) of the Criminal Code of Canada', title: 'Child Pornography — Making or Distribution',
      indictable: true, maxIndictable: '10 years imprisonment',
      irpa: 's.36(1) — serious criminality',
      amendmentLabel: 'An Act to amend the Criminal Code (child pornography), S.C. 2015, c. 23',
    },
    { fromDate: '2015-07-17', toDate: null,
      statute: 'Section 163.1(2) and (3) of the Criminal Code of Canada', title: 'Child Pornography — Making or Distribution',
      indictable: true, maxIndictable: '14 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],

  'Child Pornography / CSAM — Possession or Accessing': [
    { fromDate: '1900-01-01', toDate: '2015-07-17',
      statute: 'Section 163.1(4) and (4.1) of the Criminal Code of Canada', title: 'Child Pornography — Possession or Accessing',
      hybrid: true, maxSummary: '18 months imprisonment', maxIndictable: '5 years imprisonment',
      irpa: 's.36(2) — criminality (not serious criminality)',
      irpaThresholdCrossed: true,
      amendmentLabel: 'An Act to amend the Criminal Code (child pornography), S.C. 2015, c. 23',
    },
    { fromDate: '2015-07-17', toDate: null,
      statute: 'Section 163.1(4) and (4.1) of the Criminal Code of Canada', title: 'Child Pornography — Possession or Accessing',
      hybrid: true, maxSummary: '2 years less a day imprisonment', maxIndictable: '10 years imprisonment',
      irpa: 's.36(1) — serious criminality',
    },
  ],
};

// ─── formatMaxPunishment: builds human-readable string from a version object ───
function formatMaxPunishment(v) {
  if (v.hybrid) {
    const parts = [];
    if (v.maxSummary)   parts.push(`summary conviction: ${v.maxSummary}`);
    if (v.maxIndictable) parts.push(`indictable offence: ${v.maxIndictable}`);
    return `Hybrid offence — ${parts.join(' / ')}`;
  }
  if (v.indictable) return `Indictable offence — max. ${v.maxIndictable}`;
  if (v.summary)    return `Summary conviction — max. ${v.maxSummary}`;
  return '';
}

// ─── formatCanadianEquivalent: produces the letter-ready Canadian equivalent string ───
// Format: S. 253 "Operation While Impaired", a hybrid offence punishable by a maximum of five years' imprisonment
function formatCanadianEquivalent(v) {
  const NUM_WORDS = { '2':'two','5':'five','10':'ten','14':'fourteen','18':'eighteen','25':'twenty-five' };
  function legalMax(s) {
    if (!s) return '';
    // "5 years imprisonment" → "five years' imprisonment"
    // "18 months imprisonment" → "eighteen months' imprisonment"
    const m = s.match(/^(\d+)\s+(years|months)\s+imprisonment$/i);
    if (m) {
      const word = NUM_WORDS[m[1]] || m[1];
      return `${word} ${m[2].toLowerCase()}' imprisonment`;
    }
    if (/^2 years less a day/i.test(s)) return "two years less a day's imprisonment";
    if (/^life/i.test(s)) return 'life imprisonment';
    return s.toLowerCase();
  }
  const sMatch = v.statute.match(/[Ss]ection\s+([\w.()\-]+)/);
  const sNum   = sMatch ? sMatch[1] : v.statute;
  const type   = v.hybrid ? 'a hybrid offence' : v.indictable ? 'an indictable offence' : 'a summary conviction offence';
  const maxRaw = v.maxIndictable || v.maxSummary || '';
  const maxStr = maxRaw ? ` punishable by a maximum of ${legalMax(maxRaw)}` : '';
  return `S. ${sNum} "${v.title}", ${type}${maxStr}`;
}

// ─── lookupEquivalent: date-aware lookup using STATUTE_HISTORY ───
// Implements Tran (2010 SCC 58): use the statute in force at date of commission.
function lookupEquivalent(offenceName, offenceDate) {
  const history = STATUTE_HISTORY[offenceName];
  if (!history) return null;

  let version;
  if (offenceDate) {
    const dateStr = /^\d{4}$/.test(String(offenceDate).trim())
      ? `${offenceDate}-06-15`
      : offenceDate;
    const d = new Date(dateStr);
    version = history.find(v => {
      const from = new Date(v.fromDate);
      const to   = v.toDate ? new Date(v.toDate) : new Date('2999-12-31');
      return d >= from && d <= to;
    });
  }
  version = version || history[history.length - 1]; // fallback: current

  return {
    statute:                version.statute,
    maxPunishment:          formatMaxPunishment(version),
    canadianEquivalentLabel: formatCanadianEquivalent(version),
    irpa:                   version.irpa || '',
    isHistorical:           version.toDate !== null,
    currentStatute:         history[history.length - 1].statute,
    version,
  };
}

// ─── buildTranNote: generates the Tran paragraph for any offence where the
//     applicable statute at time of commission differs from the current statute.
//     Fires for any offence in STATUTE_HISTORY that has multiple versions.
//     Uses amendmentLabel / toDate from the applicable version for accurate citations.
function buildTranNote(offences) {
  const notes = [];
  for (const o of offences) {
    const history = STATUTE_HISTORY[o.description];
    if (!history || history.length <= 1) continue;
    if (!o.date) continue;

    const dateStr0 = /^\d{4}$/.test(String(o.date).trim()) ? `${o.date}-06-15` : o.date;
    const d = new Date(dateStr0);
    const applicable = history.find(v => {
      const from = new Date(v.fromDate);
      const to   = v.toDate ? new Date(v.toDate) : new Date('2999-12-31');
      return d >= from && d <= to;
    });
    const current = history[history.length - 1];
    if (!applicable || applicable.statute === current.statute) continue;

    // Skip if IRPA analysis is identical — same category and same max penalty means no meaningful Tran note
    if (applicable.irpa === current.irpa && applicable.maxIndictable === current.maxIndictable) continue;

    const dateStr = d.toLocaleDateString('en-CA', { year: 'numeric', month: 'long', day: 'numeric' });
    const offenceType = applicable.hybrid ? 'a hybrid offence' : 'an indictable offence';

    // Amendment label and date — defaults to Bill C-46 / Dec 18 2018 for driving offences
    const amendmentLabel = applicable.amendmentLabel || 'Bill C-46';
    const amendmentDate  = applicable.toDate
      ? new Date(applicable.toDate).toLocaleDateString('en-CA', { year: 'numeric', month: 'long', day: 'numeric' })
      : 'December 18, 2018';

    // Conclusion differs depending on whether the change crossed the IRPA s.36(1)/s.36(2) threshold
    const irpaConclusion = applicable.irpaThresholdCrossed
      ? `Accordingly, this offence constitutes criminality pursuant to section 36(2) of the IRPA, and does not rise to the level of serious criminality pursuant to section 36(1) of the IRPA.`
      : `Accordingly, while this offence constitutes serious criminality under both the former and current provisions, the applicable maximum term of imprisonment at the time of commission was ${applicable.maxIndictable}, not ${current.maxIndictable} as prescribed by the current provision. It is therefore the former statute and its corresponding maximum that govern for the purposes of this application.`;

    notes.push(
      `We note that the offence of "${o.description}" was committed on ${dateStr}, which predates the coming into force of ${amendmentLabel} on ${amendmentDate}. ` +
      `As confirmed by the Supreme Court of Canada in Canada (Citizenship and Immigration) v. Tran, 2010 SCC 58, ` +
      `for the purposes of the Immigration and Refugee Protection Act, the Canadian equivalent of a foreign offence must be assessed based on the Criminal Code as it existed at the time of commission of the foreign offence. ` +
      `Prior to ${amendmentDate}, the applicable Canadian equivalent was ${applicable.statute}, titled "${applicable.title}", ` +
      `${offenceType} punishable by a maximum term of imprisonment of ${applicable.maxIndictable}. ` +
      `${current.statute}, which came into force on ${amendmentDate}, carries a maximum of ${current.maxIndictable}. ` +
      `As the current provision was not in force at the time of our client's offence, it is the former statute that governs for the purposes of this application. ` +
      irpaConclusion
    );
  }
  return notes.join('\n\n');
}

// ─────────────────────────────────────────────
//  STATUTE REFERENCE — common US state criminal
//  statutes injected into the AI prompt so it can
//  resolve charge names ↔ statute numbers
// ─────────────────────────────────────────────
const STATUTE_REFERENCE = `
STATUTE REFERENCE TABLE — all 50 US states. Use this to resolve charge names ↔ statute numbers.
When a document says "3rd Degree DWI" or "DUI" without a statute, look up the state below and fill in the correct statute number.

ALABAMA (AL): Title 32-5A-191 = DUI; Title 13A-6-20 = Assault 1st; 13A-6-21 = Assault 2nd; 13A-6-22 = Assault 3rd; 13A-8-2 = Theft 1st; 13A-8-3 = Theft 2nd; 13A-12-212 = Possession Controlled Substance
ALASKA (AK): AS 28.35.030 = DUI; AS 11.41.200 = Assault 1st; 11.41.210 = Assault 2nd; 11.41.220 = Assault 3rd; 11.46.120 = Theft 1st; 11.71.040 = Misconduct Controlled Substance 4th
ARIZONA (AZ): ARS 28-1381 = DUI; 28-1382 = Extreme DUI (BAC 0.15+); 28-1383 = Aggravated DUI; ARS 13-1203 = Assault; 13-1204 = Aggravated Assault; 13-1802 = Theft; 13-3408 = Drug Possession/Sale
ARKANSAS (AR): ACA 5-65-103 = DWI; ACA 5-13-201 = Battery 1st; 5-13-202 = Battery 2nd; 5-13-203 = Battery 3rd; 5-36-103 = Theft; 5-64-419 = Possession Controlled Substance
CALIFORNIA (CA): VC 23152(a) = DUI Alcohol; VC 23152(b) = DUI BAC 0.08+; VC 23153 = DUI Causing Injury; VC 23550 = Felony DUI 4th offense; PC 187 = Murder; PC 211 = Robbery; PC 245 = Assault with Deadly Weapon; PC 242 = Battery; PC 261 = Rape; PC 459 = Burglary; PC 484/488 = Petty Theft; PC 487 = Grand Theft; HS 11350 = Possession Controlled Substance; HS 11351 = Possession for Sale; HS 11352 = Transport/Sale Controlled Substance; PC 496 = Receiving Stolen Property; PC 470 = Forgery
COLORADO (CO): CRS 42-4-1301 = DUI; 42-4-1301(1)(b) = DWAI; CRS 18-3-202 = Assault 1st; 18-3-203 = Assault 2nd; 18-3-204 = Assault 3rd; 18-4-401 = Theft; 18-18-403.5 = Possession Controlled Substance
CONNECTICUT (CT): CGS 14-227a = DUI/OUI; CGS 53a-59 = Assault 1st; 53a-60 = Assault 2nd; 53a-61 = Assault 3rd; 53a-119 = Larceny; 21a-279 = Possession Controlled Substance
DELAWARE (DE): 21 Del.C. 4177 = DUI; 11 Del.C. 611 = Assault 1st; 612 = Assault 2nd; 613 = Assault 3rd; 11 Del.C. 841 = Theft; 16 Del.C. 4764 = Possession Controlled Substance
FLORIDA (FL): F.S. 316.193 = DUI; F.S. 782.04 = Murder; F.S. 784.021 = Aggravated Assault; F.S. 784.03 = Battery; F.S. 784.045 = Aggravated Battery; F.S. 810.02 = Burglary; F.S. 812.014 = Theft; F.S. 893.13 = Possession/Sale Controlled Substance
GEORGIA (GA): OCGA 40-6-391 = DUI; OCGA 16-5-20 = Simple Assault; 16-5-21 = Aggravated Assault; 16-5-23 = Simple Battery; 16-5-24 = Aggravated Battery; 16-8-2 = Theft by Taking; 16-13-30 = Possession Controlled Substance
HAWAII (HI): HRS 291E-61 = OUI; HRS 707-710 = Assault 1st; 707-711 = Assault 2nd; 707-712 = Assault 3rd; HRS 708-830 = Theft; 712-1243 = Promotion of Dangerous Drug 3rd
IDAHO (ID): IC 18-8004 = DUI; IC 18-903 = Assault; IC 18-905 = Aggravated Assault; IC 18-907 = Aggravated Battery; IC 18-2403 = Theft; IC 37-2732 = Controlled Substance Violation
ILLINOIS (IL): 625 ILCS 5/11-501 = DUI; 720 ILCS 5/9-1 = Murder 1st; 5/9-2 = Murder 2nd; 5/12-3 = Battery; 5/12-3.05 = Aggravated Battery; 5/16-1 = Theft; 720 ILCS 570/402 = Possession Controlled Substance
INDIANA (IN): IC 9-30-5-1 = OWI; IC 35-42-2-1 = Battery; IC 35-42-2-3 = Aggravated Battery; IC 35-43-4-2 = Theft; IC 35-48-4-7 = Possession Controlled Substance
IOWA (IA): Iowa Code 321J.2 = OWI; Iowa Code 708.1 = Assault; 708.2 = Aggravated Misdemeanor Assault; 714.1 = Theft; 124.401 = Controlled Substance Violation
KANSAS (KS): KSA 8-1567 = DUI; KSA 21-5413 = Battery; 21-5414 = Aggravated Battery; 21-5801 = Theft; 21-5705 = Possession Controlled Substance
KENTUCKY (KY): KRS 189A.010 = DUI; KRS 508.010 = Assault 1st; 508.020 = Assault 2nd; 508.025 = Assault 3rd; 508.030 = Assault 4th; KRS 514.030 = Theft; 218A.1415 = Possession Controlled Substance
LOUISIANA (LA): RS 14:98 = DWI; RS 14:98.1 = DWI 1st offense; RS 14:98.2 = DWI 2nd offense; RS 14:98.3 = DWI 3rd offense; RS 14:98.4 = DWI 4th offense (felony); RS 14:34 = Aggravated Battery; RS 14:35 = Simple Battery; RS 14:67 = Theft; RS 40:967 = Possession Controlled Substance
MAINE (ME): 29-A MRSA 2411 = OUI; 17-A MRSA 207 = Assault; 17-A MRSA 208 = Aggravated Assault; 17-A MRSA 353 = Theft; 17-A MRSA 1107 = Possession Controlled Substance
MARYLAND (MD): TR 21-902 = DUI/DWI; CR 3-201 = Assault 1st; CR 3-203 = Assault 2nd; CR 7-104 = Theft; CR 5-601 = Possession Controlled Substance
MASSACHUSETTS (MA): MGL c.90 s.24 = OUI; MGL c.265 s.13A = Assault and Battery; c.265 s.15A = Assault and Battery with Dangerous Weapon; c.266 s.30 = Larceny; c.94C s.34 = Possession Controlled Substance
MICHIGAN (MI): MCL 257.625 = OWI/DUI; MCL 750.83 = Assault with Intent to Murder; 750.84 = Assault with Intent to do Great Bodily Harm; 750.81 = Assault and Battery; 750.81a = Aggravated Assault; MCL 750.356 = Larceny; 333.7403 = Possession Controlled Substance
MINNESOTA (MN): 169A.20 = DWI base offense; 169A.24 = 1st Degree DWI (felony); 169A.25 = 2nd Degree DWI; 169A.26 = 3rd Degree DWI; 169A.27 = 4th Degree DWI (misdemeanor); 609.185 = Murder 1st; 609.19 = Murder 2nd; 609.221 = Assault 1st; 609.222 = Assault 2nd; 609.223 = Assault 3rd; 609.2231 = Assault 4th; 609.224 = Assault 5th (misdemeanor); 609.342 = Criminal Sexual Conduct 1st; 609.343 = Criminal Sexual Conduct 2nd; 609.344 = Criminal Sexual Conduct 3rd; 609.345 = Criminal Sexual Conduct 4th; 609.52 = Theft; 609.582 = Burglary; 152.021 = Controlled Substance 1st; 152.022 = Controlled Substance 2nd; 152.023 = Controlled Substance 3rd; 152.024 = Controlled Substance 4th; 152.025 = Controlled Substance 5th; 609.595 = Criminal Damage to Property; 609.63 = Forgery
MISSISSIPPI (MS): Miss. Code 63-11-30 = DUI; 97-3-7 = Simple/Aggravated Assault; 97-17-41 = Grand Larceny; 97-17-43 = Petty Larceny; 41-29-139 = Possession Controlled Substance
MISSOURI (MO): RSMo 577.010 = DWI; RSMo 565.050 = Assault 1st; 565.052 = Assault 2nd; 565.054 = Assault 3rd; 570.030 = Stealing/Theft; 195.202 = Possession Controlled Substance
MONTANA (MT): MCA 61-8-401 = DUI; MCA 45-5-201 = Assault; 45-5-202 = Aggravated Assault; 45-6-301 = Theft; 45-9-102 = Criminal Possession Dangerous Drug
NEBRASKA (NE): NRS 60-6,196 = DUI; NRS 28-308 = Assault 1st; 28-309 = Assault 2nd; 28-310 = Assault 3rd; 28-511 = Theft; 28-416 = Possession Controlled Substance
NEVADA (NV): NRS 484C.110 = DUI; NRS 200.471 = Assault; 200.481 = Battery; 200.400 = Battery with Deadly Weapon; NRS 205.220 = Larceny; 453.336 = Possession Controlled Substance
NEW HAMPSHIRE (NH): RSA 265-A:2 = DUI; RSA 631:1 = First Degree Assault; 631:2 = Second Degree Assault; 631:2-a = Simple Assault; RSA 637:3 = Theft; 318-B:26 = Possession Controlled Substance
NEW JERSEY (NJ): NJSA 39:4-50 = DWI; NJSA 2C:12-1 = Assault; 2C:12-1(b) = Aggravated Assault; 2C:20-3 = Theft; 2C:35-10 = Possession Controlled Substance
NEW MEXICO (NM): NMSA 66-8-102 = DWI; NMSA 30-3-1 = Assault; 30-3-2 = Aggravated Assault; 30-16-1 = Larceny; 30-31-23 = Possession Controlled Substance
NEW YORK (NY): VTL 1192.1 = DWAI (Alcohol); VTL 1192.2 = DWI per se; VTL 1192.3 = DWI common law; VTL 1192.2-a = Aggravated DWI (BAC 0.18+); PL 125.25 = Murder 2nd; PL 120.10 = Assault 1st; PL 120.05 = Assault 2nd; PL 120.00 = Assault 3rd; PL 160.15 = Robbery 1st; PL 155.30 = Grand Larceny 4th; PL 220.06 = Criminal Possession CS 5th; PL 220.16 = Criminal Possession CS 3rd
NORTH CAROLINA (NC): NCGS 20-138.1 = DWI; NCGS 14-32 = Assault with Deadly Weapon; 14-33 = Simple Assault/Battery; 14-72 = Larceny; 90-95 = Possession Controlled Substance
NORTH DAKOTA (ND): NDCC 39-08-01 = DUI; NDCC 12.1-17-01 = Simple Assault; 12.1-17-02 = Aggravated Assault; 12.1-23-02 = Theft; 19-03.1-23 = Possession Controlled Substance
OHIO (OH): ORC 4511.19 = OVI/DUI; ORC 2903.01 = Aggravated Murder; 2903.02 = Murder; 2903.11 = Felonious Assault; 2903.12 = Aggravated Assault; 2903.13 = Assault; 2911.01 = Aggravated Robbery; 2911.02 = Robbery; 2913.02 = Theft; 2925.11 = Possession Controlled Substance
OKLAHOMA (OK): 47 OS 11-902 = DUI; 21 OS 641 = Assault; 21 OS 645 = Aggravated Assault and Battery; 21 OS 1704 = Larceny; 63 OS 2-402 = Possession Controlled Substance
OREGON (OR): ORS 813.010 = DUII; ORS 163.185 = Assault 1st; 163.175 = Assault 2nd; 163.165 = Assault 3rd; 164.015 = Theft; 475.894 = Possession Controlled Substance
PENNSYLVANIA (PA): 75 Pa.C.S. 3802 = DUI; 18 Pa.C.S. 2702 = Aggravated Assault; 18 Pa.C.S. 2701 = Simple Assault; 18 Pa.C.S. 3921 = Theft by Unlawful Taking; 18 Pa.C.S. 3925 = Receiving Stolen Property; 35 P.S. 780-113 = Possession Controlled Substance
RHODE ISLAND (RI): RIGL 31-27-2 = DUI; RIGL 11-5-2 = Assault; 11-5-3 = Simple Assault/Battery; 11-41-1 = Larceny; 21-28-4.01 = Possession Controlled Substance
SOUTH CAROLINA (SC): SC Code 56-5-2930 = DUI; 56-5-2933 = DUAC; SC Code 16-3-600 = Assault and Battery; 16-3-610 = Assault and Battery of High and Aggravated Nature; 16-13-30 = Petit Larceny; 16-13-40 = Grand Larceny; 44-53-370 = Possession Controlled Substance
SOUTH DAKOTA (SD): SDCL 32-23-1 = DUI; SDCL 22-18-1 = Simple Assault; 22-18-1.1 = Aggravated Assault; 22-30A-1 = Theft; 22-42-5 = Possession Controlled Substance
TENNESSEE (TN): TCA 55-10-401 = DUI; TCA 39-13-101 = Assault; 39-13-102 = Aggravated Assault; 39-14-103 = Theft; 39-17-418 = Possession Controlled Substance
TEXAS (TX): TPC 49.04 = DWI; TPC 49.045 = DWI with Child Passenger; TPC 49.07 = Intoxication Assault; TPC 49.08 = Intoxication Manslaughter; TPC 22.01 = Assault; TPC 22.02 = Aggravated Assault; TPC 22.011 = Sexual Assault; TPC 31.03 = Theft; TPC 29.02 = Robbery; TPC 29.03 = Aggravated Robbery; HSC 481.115 = Possession Controlled Substance; HSC 481.112 = Manufacturing/Delivery Controlled Substance
UTAH (UT): UCA 41-6a-502 = DUI; UCA 76-5-102 = Assault; 76-5-103 = Aggravated Assault; 76-6-404 = Theft; 58-37-8 = Possession Controlled Substance
VERMONT (VT): 23 VSA 1201 = DUI; 13 VSA 1023 = Assault; 13 VSA 1024 = Aggravated Assault; 13 VSA 2501 = Larceny; 18 VSA 4230 = Possession Controlled Substance
VIRGINIA (VA): Va. Code 18.2-266 = DUI; 18.2-57 = Assault and Battery; 18.2-51 = Malicious Wounding; 18.2-95 = Grand Larceny; 18.2-96 = Petit Larceny; 18.2-250 = Possession Controlled Substance
WASHINGTON (WA): RCW 46.61.502 = DUI; RCW 9A.36.011 = Assault 1st; 9A.36.021 = Assault 2nd; 9A.36.031 = Assault 3rd; 9A.36.041 = Assault 4th; RCW 9A.56.020 = Theft 1st; 9A.56.030 = Theft 2nd; 69.50.4013 = Possession Controlled Substance
WEST VIRGINIA (WV): WVC 17C-5-2 = DUI; WVC 61-2-9 = Assault/Battery; 61-2-10 = Malicious Wounding; 61-3-13 = Petit Larceny; 61-3-12 = Grand Larceny; 60A-4-401 = Possession Controlled Substance
WISCONSIN (WI): Wis. Stat. 346.63 = OWI/DUI; 940.19 = Battery; 940.195 = Battery to Unborn Child; 940.20 = Battery by Prisoner; 943.20 = Theft; 961.41 = Possession Controlled Substance
WYOMING (WY): WS 31-5-233 = DUI; WS 6-2-501 = Simple Assault/Battery; 6-2-502 = Aggravated Assault and Battery; 6-3-402 = Larceny/Theft; 35-7-1031 = Possession Controlled Substance

CHARGE NAME → STATUTE INFERENCE RULES:
- "DWI" / "DUI" / "OWI" / "OVI" / "DUII" / "OUI" → look up the state's impaired driving statute above
- "3rd Degree DWI" Minnesota = 169A.26 | "4th Degree DWI" Minnesota = 169A.27 | "2nd Degree DWI" Minnesota = 169A.25 | "1st Degree DWI" / "Felony DWI" Minnesota = 169A.24
- "DWAI" New York = VTL 1192.1 | "Aggravated DWI" New York = VTL 1192.2-a
- "Extreme DUI" Arizona = ARS 28-1382 | "Aggravated DUI" Arizona = ARS 28-1383
- When the document names a degree (e.g. "3rd Degree") but no statute, infer the statute from the state's degree table above
- Always use the most specific statute found in the document; if none, infer from charge name + state
`;

// ─────────────────────────────────────────────
//  OFFENCE MERGE — collapses duplicate entries
//  for the same criminal incident into one row
// ─────────────────────────────────────────────

// Extract the base statute family — strip sub-section numbers so
// "169A.20", "169A.20.1(1)", "169A.25" all map to "169A"
function statuteFamily(statute) {
  if (!statute) return '';
  const m = statute.match(/(\d+[A-Z]?\d*(?:\.\d+)?)/i);
  return m ? m[1].replace(/\.\d+.*$/, '') : statute.toLowerCase().slice(0, 10);
}

// Two offences are the "same incident" if:
//   (a) they share the same non-empty caseNumber, OR
//   (b) they share the same offence date AND the same statute family
function sameIncident(a, b) {
  // Primary key: case/docket number (cross-document merge)
  if (a.caseNumber && b.caseNumber &&
      a.caseNumber.trim() === b.caseNumber.trim()) return true;

  // Fallback: same date + same statute family
  if (!a.offenceDate || !b.offenceDate) return false;
  if (a.offenceDate !== b.offenceDate) return false;
  const fa = statuteFamily(a.foreignStatute);
  const fb = statuteFamily(b.foreignStatute);
  // if either has no statute, just match on date (conservative)
  if (!fa || !fb) return true;
  return fa === fb;
}

function mergeOffences(list) {
  const merged = [];

  for (const o of list) {
    const existing = merged.find(m => sameIncident(m, o));
    if (!existing) {
      merged.push({ ...o });
    } else {
      // Combine sentence details — avoid duplicating identical text
      const parts = [existing.sentence, o.sentence]
        .map(s => (s || '').trim())
        .filter(Boolean)
        .filter((s, i, arr) => arr.indexOf(s) === i); // dedupe exact strings
      existing.sentence = parts.join('; ');

      // Keep the most complete foreign statute (longer = more specific)
      if ((o.foreignStatute || '').length > (existing.foreignStatute || '').length) {
        existing.foreignStatute = o.foreignStatute;
      }

      // Keep latest sentence completion date
      if (o.sentenceCompletionDate && (!existing.sentenceCompletionDate ||
          o.sentenceCompletionDate > existing.sentenceCompletionDate)) {
        existing.sentenceCompletionDate = o.sentenceCompletionDate;
      }

      // Keep earliest offence date
      if (o.offenceDate && o.offenceDate < existing.offenceDate) {
        existing.offenceDate = o.offenceDate;
      }

      // Preserve case number if one side has it
      if (!existing.caseNumber && o.caseNumber) {
        existing.caseNumber = o.caseNumber;
      }
    }
  }

  return merged;
}

// ─────────────────────────────────────────────
//  LETTER TEMPLATES — style/tone/structure guides
//  keyed by travel purpose dropdown value
// ─────────────────────────────────────────────
const LETTER_TEMPLATES = {

'Cruise': `Dear Sir/Madam,

We represent the above-captioned individual in all matters pertaining to his application for a Temporary Resident Permit (TRP). An authorization indicating as much may be found enclosed herewith.

Please be advised that our client has also retained us for a Criminal Rehabilitation application, which we will be submitting to the Canadian Consulate's office in New York. As current processing times for Criminal Rehabilitation applications are estimated between twelve and eighteen months, our client is therefore requesting a Temporary Resident Permit so that he may travel to Canada in the meantime.

Reason for Request

As our client is criminally inadmissible to Canada, he is applying for a Temporary Resident Permit (TRP). Despite being classified as a rehabilitated individual, our client is applying for a TRP for added assurance during his travels.

As outlined in his travel statement and booking confirmations, our client plans to fly to Anchorage, Alaska, join a tour group, and travel on land before boarding a cruise with a port stop in British Columbia. The cruise will conclude at the port of Vancouver. Once disembarked from the ship, our client will fly out of the Vancouver airport back to the United States.

Additionally, the brief time that he and his family will be spending in Vancouver will contribute to the Canadian economy through expenditure on food, transportation, and accommodations. Our client has not had an offence in many years. Furthermore, given the brief duration of his intended stay and the transit nature of his entry into Canada, he is a low-risk traveler.

Low-Risk Traveller

We respectfully submit that the compelling reasons supporting our client's entry to Canada outweigh any minimal risk he may pose to the Canadian public. The enclosed materials and his personal statement establish that:

Accountability and Insight: He has a clear understanding of his past offences, taking full responsibility for his actions and acknowledging the seriousness of his conviction.

Remorse and Rehabilitation: Our client offers his remorse and regret for these offences. This remorse is reflected in his rehabilitated lifestyle and choices since the time of the offences.

Current Good Standing: He is now a stable, law-abiding citizen, well-established both personally and professionally. His current circumstances illustrate that he has reformed and integrated successfully into his community, with no risk factors that would undermine his rehabilitation.

Positive Reputation: Numerous enclosed reference letters attest to his good moral character, integrity, and contributions to his community.

Conclusion

Our client is not part of the prescribed class of people who pose a danger to the public. The purpose of our client's travel is brief and transit-related — he is simply passing through Canada as part of a pre-booked cruise itinerary. For all of the above reasons, we respectfully request that our client be granted a single-entry Temporary Resident Permit. We thank you for your kind consideration of this application.`,

'Wedding/Engagement': `Dear Sir/Madam,

We represent the above-captioned individual in all matters pertaining to his application for a Temporary Resident Permit. An authorization indicating as much may be found enclosed herewith.

Reason for Travel

Our client is humbly requesting a Temporary Resident Permit for deeply personal reasons. Our client is currently engaged to a Canadian citizen residing in Canada, and they plan to marry in Canada. Our client respectfully requests expedited processing of his application, as he and his partner are in the process of finalizing their wedding arrangements and must confirm details with service providers and invited guests.

This wedding represents one of the most meaningful and significant milestones in the couple's lives, and our client respectfully hopes that a past mistake will not prevent him from taking this important next step with his partner. Accordingly, our client respectfully requests the issuance of a Temporary Resident Permit to ensure he is able to enter Canada for his wedding, which is the primary and essential purpose of this request. His presence in Canada for this event is of utmost importance and cannot be deferred.

Economic Benefit to Canada

Canada's wedding industry generates billions of dollars annually, contributing meaningfully to the hospitality, catering, floral, photography, and event management sectors. Allowing our client to participate in this celebration will result in direct expenditure on accommodations, dining, venue services, and local tourism, all of which benefit the Canadian economy.

Rehabilitation Factors

Our client has demonstrated substantial personal growth since his conviction. He completed all terms of his sentence as required, has maintained a law-abiding lifestyle, and has re-established himself as a productive member of his community. The passage of time, combined with the absence of any subsequent criminal activity, strongly supports the conclusion that he no longer poses any risk to Canadian society. His motivation to attend his own wedding in Canada speaks to the genuine, positive direction his life has taken.`,

'Business Travel': `Dear Sir/Madam,

We represent the above-captioned individual in all matters pertaining to his application for a Temporary Resident Permit. An authorization indicating as much may be found enclosed herewith.

Criminal Offences:

We would like to begin by addressing the issue of our client's criminality. Enclosed is a table outlining his convictions impacting his admissibility to Canada, along with his sentencing, and the relevant foreign and Canadian statutes, for ease of reference. Per the enclosed documents, our client completed the conditions of his sentence, when his probation was terminated.

Reason for Travel

Our client is humbly requesting a Temporary Resident Permit for professional reasons. Our client is employed in a senior capacity and will be travelling to Canada to fulfill important business obligations that cannot be completed remotely.

As a foreign national engaged in international business activities without any intention of directly entering the Canadian labour market, our client, as per R187.1 of the Immigration and Refugee Protection Regulations, does not require a work permit.

Furthermore, the Immigration and Refugee Protection Regulations create a class of foreign nationals who are eligible to work without needing a work permit as "business visitors" who seek to engage in international business activities in Canada without directly entering the Canadian labour market, provided the primary source of remuneration for the business activities is outside Canada and the principal place of business and actual place of accrual of profits remain predominantly outside Canada.

Economic Benefit to Canada

During his stay, our client will contribute to the local economy through dining, accommodations, shopping, and the use of local transportation services. More broadly, his business activities in Canada will foster meaningful professional collaborations with Canadian entrepreneurs and businesses, thereby benefiting both Canadians and the local economy.

Canada's business tourism sector generates billions in economic activity annually. Hosting international business visitors supports high-skilled employment in the hospitality, transportation, and professional services sectors and reinforces Canada's reputation as a world-class destination for commerce and investment.

Rehabilitation Factors

Our client has demonstrated a sustained commitment to lawful conduct since his conviction. He completed all sentencing requirements in full, has maintained steady employment, and has shown no subsequent criminal activity. The conviction represents an isolated incident from which he has clearly moved on. His professional standing and ongoing career advancement are a testament to his rehabilitation and good character.`,

'Business/Conference/Event': `Dear Sir/Madam,

We represent the above-captioned individual in all matters pertaining to his application for a Temporary Resident Permit. An authorization indicating as much may be found enclosed herewith.

Criminal Offences:

We would like to begin by addressing the issue of our client's criminality. Enclosed is a table outlining his convictions impacting his admissibility to Canada, along with his sentencing, and the relevant foreign and Canadian statutes, for ease of reference. Per the enclosed documents, our client completed all conditions of his sentence.

Reason for Travel

Our client respectfully requests the issuance of a Temporary Resident Permit for professional reasons. We acknowledge that Temporary Resident Permits are issued only in exceptional circumstances and respectfully submit that this matter constitutes such a case. The purpose of our client's travel is strictly business-related and tied to his employment.

Our client intends to travel to Canada to attend a professional conference or industry event, bringing together leaders to share knowledge, engage in discussions on emerging trends, and foster business development through workshops, educational sessions, and networking events. These opportunities allow professionals to establish and strengthen relationships, thereby fostering collaboration among attendees from around the world.

Given his title and experience, our client holds critical knowledge and leadership authority that cannot be delegated or replicated remotely for this event.

Furthermore, as a foreign national engaged in international business activities without any intention of directly entering the Canadian labour market, our client, as per R187.1 of the Immigration and Refugee Protection Regulations, does not require a work permit.

Economic Benefit to Canada

During his stay, our client will reside in a Canadian hotel and plans to contribute to the local economy through dining, shopping, and the use of local transportation services. More broadly, his business activities in Canada will foster meaningful professional collaborations with Canadian entrepreneurs and businesses, thereby benefiting both Canadians and the local economy.

Canada's conference and events industry is a significant contributor to the national economy, generating revenue through accommodations, food services, transportation, and ancillary tourism spending. International business visitors like our client bring foreign investment and professional exchange that strengthen Canada's position as a global hub for industry and innovation.`,

'Family Visit': `Dear Sir/Madam,

We represent the above-captioned individual in all matters pertaining to his application for a Temporary Resident Permit (TRP). An authorization indicating as much may be found enclosed herewith.

Reason for Request

As our client is criminally inadmissible to Canada, he is applying for a Temporary Residence Permit. He is requesting a single-entry to Canada with the purpose of reuniting with his family.

Family reunification is a core principle guiding the policies of Immigration, Refugees and Citizenship Canada (IRCC), reflecting Canada's commitment to compassion, stability, and social cohesion. Facilitating the applicant's temporary entry would support the humanitarian objective of preserving family unity and allowing meaningful participation in an important family gathering.

Milestone events carry profound emotional importance and provide a rare opportunity for families to come together, particularly when separation has been prolonged due to immigration barriers. Given the exceptional nature of this event, granting a TRP would help prevent undue hardship to both the applicant and his family members, while remaining consistent with the spirit and intent of Canada's immigration policies.

Economic Benefit to Canada

Tourism is a significant contributor to Canada's economy, generating substantial revenue, employment, and tax receipts across multiple sectors. In 2024, tourism activities generated approximately $129.7 billion in revenue distributed across services such as accommodation, food and beverage, transportation, recreation, and retail, surpassing pre-pandemic levels and demonstrating strong continued growth in the travel sector.

Allowing our client temporary permission to enter Canada is in the country's economic interest because he will actively engage with the tourism and hospitality sectors throughout his visit. He intends to dine, shop, and participate in various cultural and recreational activities, making meaningful financial contributions to local and national businesses.

Rehabilitation Factors

Our client has demonstrated a sustained commitment to lawful conduct in the years since his conviction. He completed all sentencing requirements as imposed by the court. The conviction represents an isolated incident that does not define who he is today. His desire to reunite with family in Canada reflects the genuine, positive values that now guide his life.`,

'Partner/Spouse Lives in Canada': `Dear Sir/Madam,

We represent the above-captioned individual in all matters pertaining to his application for a Temporary Resident Permit (TRP). An authorization indicating as much may be found enclosed herewith.

Reason for Request:

As our client is currently criminally inadmissible to Canada, he is applying for a Temporary Resident Permit (TRP). Our client has a confirmed trip to Canada with his partner, a Canadian citizen whose family resides in Canada.

Our client is applying for a one-year multiple-entry TRP to allow the couple the opportunity to visit Canada on multiple occasions throughout the year, including family gatherings, holidays, and important milestones such as weddings and celebrations planned by his partner's family. His partner has been working abroad and has not returned home in several years. If the TRP is granted, they intend to travel back to Canada to spend meaningful time with her family.

Benefit to Canada:

Granting our client a TRP will provide tangible economic benefits to Canada through tourism and hospitality spending. During their visit, our client and his partner will stay at hotels, dine at local restaurants, shop in various cities, rent a vehicle for transportation, and purchase tickets for local attractions. These activities will directly support local businesses and contribute to the Canadian economy across multiple sectors, including accommodations, dining, retail, and transportation.

It is also important to note that tourism is a significant contributor to Canada's economy, generating substantial revenue, employment, and tax receipts across multiple sectors. In 2024, tourism activities generated approximately $129.7 billion in revenue across services such as accommodation, food and beverage, transportation, recreation, and retail, surpassing pre-pandemic levels.

Allowing our client temporary permission to enter Canada is in the country's economic interest because he will actively engage with the tourism and hospitality sectors throughout his visit.

Rehabilitation Factors

Our client has multiple convictions, all of which have been disclosed and addressed herein. Enclosed herewith is a detailed table of offences that describes our client's convictions alongside their appropriate Canadian equivalencies. Despite this history, our client has demonstrated genuine rehabilitation through his stable personal life, continued employment, and absence of any further criminal activity. His long-term relationship with a Canadian citizen further demonstrates the ties and responsibilities that anchor him to a lawful, productive lifestyle.`,

'Tourism/Leisure': `Dear Sir/Madam,

We represent the above-captioned individual in all matters pertaining to his application for a Temporary Resident Permit (TRP). An authorization indicating as much may be found enclosed herewith.

Reason for Request

As our client is criminally inadmissible to Canada, he is applying for a Temporary Residence Permit.

Our client is planning a leisure trip to Canada — a recreational visit centered around outdoor activities in a stunning natural setting. The purpose of the visit is leisure and family connection, consisting of organized recreational activities with family members who share a longstanding tradition of visiting Canada together.

This trip will be undertaken with family members, with whom our client has a longstanding tradition of enjoying Canadian destinations. While the primary purpose of travel is recreational, it also reflects the broader principles of family reunification recognized by Immigration, Refugees and Citizenship Canada. Canadian immigration policy places importance on enabling families to spend time together, even in a temporary context, as a means of maintaining and strengthening close familial relationships. This aligns with the intent of temporary resident travel, combining leisure with meaningful family connection, while maintaining a clear and limited duration of stay and full intention to depart Canada at the end of the visit.

Benefit to Canada:

Our client and his group plan to contribute substantially to Canada's tourism and adventure tourism industries during their trip. The upfront expense of this trip has already been considerable, representing a meaningful contribution to the Canadian economy. They plan to enjoy local activities and intend on dining out and shopping frequently at local establishments during their stay.

Canada's leisure tourism industry generates substantial economic activity, benefiting local businesses and service providers in communities across the country. Adventure tourism, outdoor recreation, and resort destinations generate hundreds of millions of dollars annually and support thousands of jobs in the accommodation, food service, retail, and guiding sectors.

Rehabilitation Factors

Our client has completed all aspects of his sentence requirements and has been discharged from all court-imposed obligations. He has maintained a clean record since the time of his conviction and has worked diligently to rebuild his life. His participation in this family trip reflects his current stable, law-abiding lifestyle and the meaningful personal relationships he has cultivated.`,

'Flight Attendant': `Dear Sir/Madam,

We represent the above-named individual with respect to her application for a Temporary Resident Permit. An authorization to this effect is enclosed.

We would like to start by addressing the issue of our client's criminality. Her singular conviction is set out below in tabular form for ease of reference, together with her sentence, the U.S. statute under which she was convicted, and the Canadian equivalent.

Please note that our client has also retained us to prepare a Criminal Rehabilitation application on her behalf, which we will submit to the Consulate in New York once she becomes eligible. We are therefore proceeding with a Temporary Resident Permit in the meantime.

Our client is humbly requesting a Temporary Resident Permit to enter Canada for professional purposes, commencing as soon as possible and for a period of one year. We assert that our client is an exemplary candidate for a Temporary Resident Permit and fully comprehends its temporary nature. Our client has no prior or subsequent legal infractions and therefore does not exhibit the characteristics of an individual who poses a threat to Canadian society.

Our client is seeking a Temporary Residence Permit for professional purposes. Our client has worked as a flight attendant and now hopes to advance her career as an international flight attendant. Several airlines have expressed interest in hiring our client for international routes. This opportunity, however, requires that she obtain a Temporary Resident Permit. As a flight attendant, our client's responsibilities include ensuring the safety and comfort of passengers on flights to and from a wide range of destinations. Given that her professional duties may require entry into Canada, it is essential that she be able to enter and depart the country without restriction. Moreover, in the event of an emergency landing, she would be legally obligated to land at any available location, which could include Canada. For these reasons, it is critical that our client be permitted entry into Canada.

The passengers that our client would accompany in and out of Canada are predominantly people contributing to the Canadian economy either through business or tourism. By ensuring comfortable travel for tourists and business persons alike, our client would be fostering the goodwill that already exists between Canada and the United States. Thus our client would be, in a direct sense, contributing to the Canadian economy by transporting citizens of both countries in this manner. Our client's presence in Canada will likewise stimulate local economies. Given her need to frequently stay in Canada, she will bring business to the tourism and hospitality sectors in several Canadian cities, including local hotels, restaurants, and transportation services. Therefore, it would be in Canada's best interest that she is able to carry out her professional duties.

We contend that our client's need to enter Canada, as described above, greatly outweighs her potential inadmissibility. As indicated in her personal statement, our client fully comprehends the repercussions of her conviction and assumes full accountability for its impact on the trajectory of her life. In addition to fulfilling her court-mandated sentence, our client has remained committed to enacting lasting positive changes. She expresses profound regret for her actions and remains remorseful. This is evident as she has had no subsequent encounters with the law. Our client maintains that she has learned a valuable lesson and has become a better person as a result.

Our client's enclosed reference letters highlight her motivation, kindness, and professionalism. She is particularly praised for her patience and adaptability, even in challenging situations. Her commitment to human connection sets her apart both professionally and personally. Her unwavering reliability has been a key factor in her successes. Although our client made a mistake in her past, those around her contend that she has worked diligently to rectify it. These letters indicate that our client's conviction is not a reflection of her true character whatsoever, but rather a mistake for which she takes full responsibility.

In addition, please be advised that as a foreign national engaged in international business activities without any intention of directly entering the Canadian labour market, our client, as per R187.1 of the Immigration and Refugee Protection Regulations, requires no work permit.

For all of the above reasons, we respectfully request that our client be provided with a Temporary Resident Permit to enter into Canada commencing as soon as possible and for a period of one year.`,

'Pilot': `Dear Sir/Madam,

We represent the above-named individual with respect to his application for a Temporary Resident Permit.

Please note that we already submitted a version of this application to the Canadian Consulate, but have not yet received a decision. We are therefore requesting that you process our client's application today at the port of entry.

We would like to address the issue of our client's criminality. Our client's conviction is set out below in tabular form for ease of reference, together with his sentence, the U.S. statute under which he was convicted, and the Canadian equivalent.

Please note that our client has also retained us to prepare a Criminal Rehabilitation application on his behalf, which we have submitted to the Consulate in New York, as confirmed in the enclosed Acknowledgment of Processing Letter. The typical waiting times for this kind of application are currently in excess of 18 months, so we are therefore proceeding with a Temporary Resident Permit in the meantime.

Our client is requesting a multiple-entry Temporary Resident Permit to enter Canada for a period of one year, beginning as soon as possible. We contend that our client is an ideal candidate for a Temporary Resident Permit and understands that it will be temporary in nature. He has had a regretful interaction with the law and, as such, does not possess the profile of someone who poses a threat to Canadian society.

Our client's presence is required in Canada for professional reasons. He is currently employed as a regional pilot by an airline which involves flights into Canadian airports on an ongoing basis. As a pilot, allowing our client to enter Canada would not only enable him to perform his professional duties but also contribute economically. His work involves transporting numerous passengers who travel for business and leisure, thus fostering tourism and business travel that benefit the Canadian economy. The nature of his job requires mobility and reliability, traits that he has demonstrated consistently throughout his career. Granting our client the ability to travel to Canada would allow him to continue making positive contributions to society and the economy. It is in the best interests of all involved parties to consider the broader implications of his professional activities, which have significant economic benefits for Canada.

We firmly believe that our client's need to enter Canada, as detailed above, far outweighs any concerns regarding his potential inadmissibility. As stated in his personal statement, our client is aware of the consequences of his past actions and takes full responsibility. He has gained a deeper understanding of his actions and the danger he posed to others. His statement reflects deep regret, and he has made a conscious effort to prioritize safe driving ever since. Our client has diligently completed all required assessments and services as part of his sentence, ensuring such incidents will not recur.

Furthermore, the letters of reference enclosed with our client's application paint a consistent picture of him as a responsible individual of outstanding character and integrity. He is described as a disciplined, law-abiding individual in possession of a strong moral compass. Our client's references explain that he has matured greatly since the time of his charges. He has since found his direction in life and is actively building a positive future. His references emphasize that while he made mistakes in the past, he has earnestly worked to make amends. These letters assert that our client's convictions do not reflect his true character in any way.

For all of the above reasons, we respectfully request that our client be provided with a multiple-entry Temporary Resident Permit valid for a period of one year and beginning as soon as possible.`,

'Hunting/Fishing': `Dear Sir/Madam,

We represent the above-captioned individual in all matters pertaining to his application for a Temporary Resident Permit (TRP). An authorization indicating as much may be found enclosed herewith.

Please note that current processing times for Temporary Resident Permit applications submitted to the Consulate are advertised at 3-6 months. As there is insufficient time for the Consulate to process it, we respectfully request that you process this application for a Temporary Resident Permit today at the Port of Entry.

According to Enforcement Manual (ENF) 4: A border services officer at Immigration Secondary has discretion, pursuant to subsection A24(1), to issue a TRP to an inadmissible person seeking entry to Canada if satisfied that entry is justified in the circumstances.

Reason for Request

As our client is criminally inadmissible to Canada, he is applying for a Temporary Residence Permit. He is requesting entry to Canada to go on a hunting trip organized by authorized local outfitters in British Columbia. We are thus humbly requesting that our client be allowed to enter Canada to enjoy the Canadian wilderness and partake in activities in the region. Our client's trip will prove to be of large economic benefit to the province of British Columbia and thus we contend that he is an ideal candidate for a Temporary Resident Permit.

As mentioned, our client will undoubtedly provide a significant contribution to the Canadian tourism economy during this trip. It is therefore not only in the best economic interest of the province of British Columbia, but Canada as a whole, to allow our client to participate in the upcoming hunting trip. The cost of this hunting trip is substantial.

Lastly, hunters directly fund conservation efforts in British Columbia. Every hunting license and species tag includes a conservation surcharge that goes to Habitat Conservation Trust Foundation (HCTF), which funds over 100 fish, wildlife and habitat projects above and beyond government funding annually.

Low-Risk Traveller

We respectfully submit that the compelling reasons for our client's entry into Canada outweigh the relatively minimal risk that he may pose to the Canadian public. It is clear from the enclosed evidence and personal statement that our client has a clear understanding of his past offences, taking full responsibility for his actions and acknowledging the seriousness of his conviction. Our client has demonstrated remorse and is living a rehabilitated lifestyle.`,

};

// ─────────────────────────────────────────────
//  ROUTES
// ─────────────────────────────────────────────

app.get('/api/equivalency-bank', (req, res) => {
  res.json(Object.keys(STATUTE_HISTORY));
});

app.get('/api/equivalency-lookup', (req, res) => {
  const { offence, date } = req.query;
  const result = lookupEquivalent(offence, date);
  res.json(result || { statute: '' });
});

// ─── PASSPORT PARSE (image or PDF) ───
app.post('/api/parse-passport', upload.single('passport'), async (req, res) => {
  if (!API_KEY) return res.status(503).json({ error: 'ANTHROPIC_API_KEY not set' });
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  try {
    let messageContent;
    const mime = req.file.mimetype;
    const extractPrompt = `You are reading a passport biographical data page. Extract the holder's information using ALL available text in the image.

IMPORTANT: Look specifically for the MRZ (Machine Readable Zone) — the two lines of text at the very bottom of the passport page that contain "<" characters (e.g. "P<USASMITH<<JOHN<MICHAEL<<<<<<<<<<<<<<<<<<<<<"). The MRZ is the most reliable source.

MRZ line 1 format: P<[country][surname]<<[given names with < separators]
MRZ line 2 format: [passport number][check][nationality][YYMMDD birth][check][sex][YYMMDD expiry][check][personal number]

Also read the printed biographical fields (Name, Date of birth, Nationality, Passport No.) visible on the page.

Return ONLY valid JSON — no other text:
{ "name": "GIVEN NAMES SURNAME (as printed, not MRZ format)", "dob": "YYYY-MM-DD", "nationality": "full country name (e.g. United States of America)", "passportNumber": "passport number without spaces" }
If a field cannot be read, use null.`;

    if (mime === 'application/pdf') {
      const data = await pdfParse(req.file.buffer);
      messageContent = [{ type: 'text', text: `${extractPrompt}\n\nPassport document text:\n${data.text.slice(0, 3000)}` }];
    } else {
      const mediaType = ['image/jpeg','image/png','image/gif','image/webp'].includes(mime) ? mime : 'image/jpeg';
      messageContent = [
        { type: 'image', source: { type: 'base64', media_type: mediaType, data: req.file.buffer.toString('base64') } },
        { type: 'text', text: extractPrompt },
      ];
    }

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'x-api-key': API_KEY, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
      // Use Sonnet for better vision/OCR accuracy on passport images
      body: JSON.stringify({ model: 'claude-sonnet-4-6', max_tokens: 512, messages: [{ role: 'user', content: messageContent }] }),
    });

    const aiData = await response.json();
    const raw = aiData.content[0].text.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
    res.json(JSON.parse(raw));
  } catch (err) {
    console.error('Passport parse error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─── PLAIN TEXT EXTRACT (travel statement, rehab docs) ───
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const DOC_MIME  = 'application/msword';

app.post('/api/extract-text', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const mime = req.file.mimetype;
    let text = '';
    if (mime === 'application/pdf') {
      const data = await pdfParse(req.file.buffer);
      text = data.text.trim();
    } else if (mime === DOCX_MIME || mime === DOC_MIME ||
               req.file.originalname.match(/\.docx?$/i)) {
      // mammoth handles both .docx (and attempts .doc)
      const result = await mammoth.extractRawText({ buffer: req.file.buffer });
      text = result.value.trim();
    } else if (mime.startsWith('text/')) {
      text = req.file.buffer.toString('utf-8');
    } else {
      return res.status(400).json({ error: 'Unsupported file type. Please upload a PDF, Word document (.docx), or text file.' });
    }
    res.json({ text: text.slice(0, 8000) });
  } catch (err) {
    console.error('Text extract error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─── PDF PARSE ───
app.post('/api/parse-pdf', upload.single('pdf'), async (req, res) => {
  if (!API_KEY) return res.status(503).json({ error: 'ANTHROPIC_API_KEY not set' });
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  try {
    const data = await pdfParse(req.file.buffer);
    const text = data.text.slice(0, 6000); // cap to avoid token limits

    console.log('\n════════ PDF TEXT EXTRACTED ════════');
    console.log(text);
    console.log('════════════════════════════════════\n');

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': API_KEY,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 2048,
        system: `You extract criminal offence information from court documents. Return ONLY valid JSON, no extra text.`,
        messages: [{
          role: 'user',
          content: `${STATUTE_REFERENCE}

Extract criminal offences from this court document. Return JSON with these fields:
- caseNumber (the court case/file/docket number exactly as it appears — e.g. "27-CR-15-12345". If multiple case numbers appear, use the first/primary one. This is CRITICAL for cross-document merging.)
- offenceDate (YYYY-MM-DD or best estimate of the date the offence occurred)
- offenceDescription (standard legal name of the charge — see statute guidance below)
- sentence (complete sentence imposed, combining all sentencing information found — e.g. "2 years probation, $1,000 fine, license revoked 1 year")
- sentenceCompletionDate (always return null — this field is filled in manually by the case manager)
- foreignStatute (exact statute cited in the document, e.g. "Minnesota Statute 169A.20.1(1)")
- country (country where offence occurred)

CONSOLIDATION RULES — very important:
- Court documents often contain multiple entries for the SAME offence (e.g. original charge, amended charge, sentencing hearing, probation discharge, licence revocation). These must be merged into ONE offence record.
- Two entries are the SAME offence if they share the same case/docket number OR the same offence date + statute
- When merging, combine all sentencing details into a single sentence string (e.g. "18 months probation; $680 fine; licence suspended 1 year; discharged 2017-03-15")
- Use the earliest/most specific offence date found across all related entries
- Use the latest date mentioned as the sentenceCompletionDate (discharge, completion, expiry date)
- The final result should have ONE entry per distinct criminal incident, not one entry per court appearance or document section

STATUTE GUIDANCE — use the statute number to determine the correct offence name:
- 169A.20 / 169A (Minnesota), VC 23152 (California), RS 14:98 (Louisiana), or any DWI/DUI/OWI/OUI statute → "DUI / DWI / Impaired Driving (Alcohol)"
- Statutes referencing "dangerous operation", "reckless driving" → "Dangerous Driving / Dangerous Operation"
- Statutes referencing "assault", "battery" → use the appropriate assault level
- Statutes referencing "theft", "larceny", "shoplifting" → "Theft Under $5,000" or "Theft Over $5,000" based on amount
- When in doubt, use the plain English charge name from the document verbatim

CRITICAL — do NOT confuse statutes:
- A DWI/DUI statute is NEVER theft, even if the word "theft" appears elsewhere in the document
- Base offenceDescription on what the statute actually criminalises, not on surrounding text

EXCLUDE — do NOT include:
- Minor traffic infractions: speeding, failure to yield/stop, improper lane change, illegal turn, tailgating, expired registration, no seatbelt, distracted driving, running a red light, rolling stop, parking violations
- Only include criminal charges (DUI/DWI, dangerous driving, assault, theft, fraud, drugs, weapons, etc.)

Document text:
${text}

Return only: { "offences": [ ... ] }`
        }],
      }),
    });

    const aiData = await response.json();
    const raw = aiData.content[0].text.trim()
      .replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');

    console.log('\n════════ AI EXTRACTED OFFENCES ════════');
    console.log(raw);
    console.log('═══════════════════════════════════════\n');

    const parsed = JSON.parse(raw);
    parsed.offences = mergeOffences(parsed.offences || []);

    console.log('\n════════ AFTER MERGE ════════');
    console.log(JSON.stringify(parsed.offences, null, 2));
    console.log('══════════════════════════════\n');

    res.json(parsed);
  } catch (err) {
    console.error('PDF parse error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─── GENERATE LETTER ───
app.post('/api/generate-letter', async (req, res) => {
  if (!API_KEY) return res.status(503).json({ error: 'ANTHROPIC_API_KEY not set' });

  const { client, travel, travelStatement, rehabilitation, attachments, applicationMode } = req.body;

  // Support both new arrests[] and legacy offences[] formats
  const body = req.body;
  const arrestsData = body.arrests || (body.offences || []).map(o => ({
    date: o.date, country: o.country,
    charges: [{ description: o.description, outcome: o.outcome, sentence: o.sentence,
                programConditions: '', foreignStatute: o.foreignStatute,
                canadianStatute: o.canadianStatute, canadianEquivalentLabel: o.canadianEquivalentLabel || '',
                irpa: o.irpa }],
    noRecords: o.noRecords, noRecordsReason: o.noRecordsReason,
  }));

  const offenceSummary = arrestsData.map((arr, i) => {
    const chargeLines = (arr.charges || []).map(c =>
      `  - ${c.description} (outcome: ${c.outcome}${c.sentence ? ', sentence: ' + c.sentence : ''}${c.programConditions ? ', program conditions: ' + c.programConditions : ''}). Foreign statute: ${c.foreignStatute}. Canadian equivalent: ${c.canadianEquivalentLabel || c.canadianStatute || 'TBD'}.`
    ).join('\n');
    return `Arrest ${i+1} on ${arr.date} in ${arr.country}:\n${chargeLines}${arr.noRecords ? `\n  [NOTE: Court records unavailable — reason: ${arr.noRecordsReason === 'untimely' ? 'client unable to obtain on time' : 'records purged'}]` : ''}`;
  }).join('\n\n');

  // Build verbatim court-record unavailability paragraphs — group by reason
  function noRecordsDateList(arrs) {
    const dates = arrs.map(a => a.date ? `(${a.date})` : '');
    if (dates.length === 1) return `the ${dates[0]} arrest`;
    if (dates.length === 2) return `the ${dates[0]} and ${dates[1]} arrests`;
    return `the ${dates.slice(0,-1).join(', ')}, and ${dates[dates.length-1]} arrests`;
  }
  const purgedArrs   = arrestsData.filter(a => a.noRecords && a.noRecordsReason !== 'untimely');
  const untimelyArrs = arrestsData.filter(a => a.noRecords && a.noRecordsReason === 'untimely');
  const noRecordsParagraphs = [];
  if (purgedArrs.length)
    noRecordsParagraphs.push(`Records relating to ${noRecordsDateList(purgedArrs)} are no longer available, having been destroyed in accordance with the court's record retention policy. Accordingly, our client has provided a signed statement setting out the sentencing details to the best of their recollection.`);
  if (untimelyArrs.length)
    noRecordsParagraphs.push(`Our client was unable to obtain the relevant court records for ${noRecordsDateList(untimelyArrs)} prior to the submission of this application. Accordingly, they have provided a signed statement setting out the relevant sentencing details to the best of their recollection.`);
  const noRecordsBlock = noRecordsParagraphs.length
    ? noRecordsParagraphs.join('\n\n')
    : null;

  // Build criminal rehabilitation paragraph if applicable
  const rehabStatus    = rehabilitation?.rehabStatus || 'none';
  const rehabAORDate   = rehabilitation?.rehabAORDate || '';
  const rehabAORNumber = rehabilitation?.rehabAORNumber || '';
  const lastName = (client.name || '').split(' ').pop();
  const gender  = client.gender || 'mr';
  const title   = gender === 'ms' ? 'Ms.'  : 'Mr.';
  const heShe   = gender === 'ms' ? 'she'  : 'he';
  const hisHer  = gender === 'ms' ? 'her'  : 'his';
  const himHer  = gender === 'ms' ? 'her'  : 'him';
  let rehabParagraph = '';
  if (rehabStatus === 'applied') {
    const dateStr = rehabAORDate
      ? new Date(rehabAORDate).toLocaleDateString('en-CA', { year:'numeric', month:'long', day:'numeric' })
      : '[DATE]';
    const fileRef = rehabAORNumber ? ` (${rehabAORNumber})` : '';
    rehabParagraph = `Please be advised that our client has also retained us for a Criminal Rehabilitation application, which we submitted to the Canadian Consulate's office in New York, as confirmed in the enclosed Acknowledgement of Receipt letter dated ${dateStr}${fileRef}. As current processing time for Criminal Rehabilitation applications is estimated between twelve and eighteen months, ${title} ${lastName} is therefore requesting a Temporary Resident Permit so that ${heShe} may travel to Canada in the meantime.`;
  } else if (rehabStatus === 'retained') {
    rehabParagraph = `Please be advised that our client has also retained us for a Criminal Rehabilitation application, which we will be submitting to the Canadian Consulate's office in New York. As current processing times for Criminal Rehabilitation applications are estimated between twelve and eighteen months, ${title} ${lastName} is therefore requesting a Temporary Resident Permit so that ${heShe} may travel to Canada in the meantime.`;
  }

  // Build POE paragraph (verbatim — not AI-generated)
  const appType = applicationMode?.type || 'consulate';
  const ENF_BLOCK =
    `According to Enforcement Manual (ENF) 4:\n\n` +
    `A border services officer at Immigration Secondary has discretion, pursuant to subsection A24(1), to issue a TRP to an inadmissible person seeking entry to Canada if satisfied that entry is justified in the circumstances.\n\n` +
    `Furthermore, IRCC has confirmed that TRPs can be processed at the ports of entry. Please find an email from IRCC dated March 27, 2017, enclosed herewith.`;

  let poeParagraph = '';
  if (appType === 'poe-sole') {
    poeParagraph =
      `Please note that current processing times for Temporary Resident Permit applications submitted to the Consulate are advertised at 3-6 months. As there is insufficient time for the Consulate to process it, we respectfully request that you process this application for a Temporary Resident Permit today at the Port of Entry.\n\n` +
      ENF_BLOCK;
  } else if (appType === 'poe-prior-consulate') {
    const city    = applicationMode.consulateCity  || '[CONSULATE CITY]';
    const dateStr = applicationMode.submissionDate
      ? new Date(applicationMode.submissionDate).toLocaleDateString('en-CA', { year: 'numeric', month: 'long', day: 'numeric' })
      : '[DATE]';
    const aorRef  = applicationMode.aorNumber
      ? `Application Number: ${applicationMode.aorNumber}`
      : 'Application Number: [NUMBER]';
    poeParagraph =
      `Please be advised that we have submitted a version of this application to the Canadian Consulate in ${city} on ${dateStr}, as indicated in the enclosed Acknowledgment of Processing Letter (${aorRef}). As we have not yet received a decision on this matter, we respectfully request that you process this application for a Temporary Resident Permit today at the Port of Entry.\n\n` +
      ENF_BLOCK;
  }

  // Look up style template for this travel purpose
  const templateText = LETTER_TEMPLATES[travel.purpose] || null;

  // Build Tran paragraph for any pre-2018 driving offences
  // Flatten to offence-like objects for buildTranNote compatibility
  const flatOffencesForTran = arrestsData.flatMap(arr =>
    (arr.charges || []).map(c => ({ description: c.description, date: arr.date }))
  );
  const tranNote = buildTranNote(flatOffencesForTran);

  const prompt = `You are drafting a Temporary Resident Permit (TRP) cover letter for Cohen Immigration Law (attorney Daniel Levy) on behalf of a criminally inadmissible foreign national seeking entry to Canada.

CLIENT: ${client.name}, DOB: ${client.dob}, Nationality: ${client.nationality || 'N/A'}
PRONOUNS: ${title} ${lastName} — use "${heShe}" / "${hisHer}" / "${himHer}" throughout
TRAVEL PURPOSE: ${travel.purpose}
ENTRY TYPE: ${travel.entryType}
PROPOSED DATES: ${travel.dateFrom} to ${travel.dateTo}
LOCATIONS IN CANADA: ${travel.locations}
DURATION: ${travel.duration}

OFFENCE HISTORY:
${offenceSummary || 'No offences provided'}

TRAVEL STATEMENT (provided by client):
${travelStatement || 'Not provided'}

REHABILITATION SUMMARY (may include content extracted from reference letters and supporting documents):
${rehabilitation.summary || 'Not provided'}
${tranNote ? `
TRAN PARAGRAPH — MANDATORY VERBATIM INSERTION:
One or more offences predate December 18, 2018 and require a Tran analysis paragraph. You MUST insert the following paragraph(s) VERBATIM, word for word, into the "Foreign Offence(s)" / introduction section of the letter — immediately after the conviction table is referenced. Do NOT paraphrase, shorten, or alter this text in any way:

${tranNote}
` : ''}${noRecordsBlock ? `
COURT RECORDS UNAVAILABLE — MANDATORY VERBATIM INSERTION:
Court records are unavailable for one or more offences. You MUST insert the following paragraph(s) VERBATIM, word for word, into the "Foreign Offence(s)" section of the letter — immediately after the conviction details are presented. Do NOT paraphrase, shorten, or alter this text in any way:

${noRecordsBlock}
` : ''}
NOTE — SENTENCE FORMAT IN OFFENCE TABLE: When referencing the sentence imposed in the letter, use point form (e.g. "Probation: 2 years; Fine: $680; Licence suspension: 1 year").
${templateText ? `
STYLE GUIDE — Reference letter for "${travel.purpose}" travel:
The following is a real Cohen Immigration Law TRP letter for a similar type of travel. Use it as a STYLE and STRUCTURE guide ONLY. Mirror its tone, vocabulary, paragraph length, legal language, and section organization. Do NOT copy client names, dates, statute numbers, or specific facts — replace all specifics with the current client's information. The goal is to match the professional voice, sentence cadence, and argumentative approach used in this type of letter.

--- REFERENCE LETTER START ---
${templateText}
--- REFERENCE LETTER END ---
` : ''}
Generate the following sections as JSON. Each section should be professional legal prose, 2-4 paragraphs each. Write as if you are the attorney.

Return ONLY this JSON structure:
{
  "introduction": "Draft 1–2 paragraphs of introduction prose that follow the standard opening sentence and any POE/rehabilitation paragraphs. Do NOT re-state 'We represent...' or repeat the standard opening — that sentence is already inserted verbatim before your text. Do NOT include the POE or criminal rehabilitation paragraphs — those are also already inserted verbatim. Focus on: briefly framing the inadmissibility, introducing the case, and leading into the numbered sections.",
  "reasonForRequest": "2-3 paragraphs explaining the purpose of the visit, travel plans, and why client needs to enter Canada. Incorporate the travel statement details naturally.",
  "economicBenefit": "2 paragraphs on economic benefit of the visit to Canada. Include specific realistic statistics about tourism/business revenue, citing real Canadian government sources. Add footnote markers like [1], [2] in the text.",
  "economicBenefitFootnotes": ["Footnote 1 text with full citation", "Footnote 2 text with full citation"],
  "rehabilitationFactors": "2-3 paragraphs on rehabilitation. Reference the personal statement, reference letters, time elapsed since offence, and positive changes in client's life.",
  "conclusion": "EXACTLY 3 SENTENCES — no more, no less. VIOLATION OF THIS RULE IS NOT PERMITTED. Do not recap, summarize, or repeat anything from prior sections. Sentence 1: The specific TRP request — entry type, applicant name, dates/duration. Sentence 2: ONE closing argument only (e.g. isolated offence, completed sentence, no subsequent criminality, compelling purpose). Sentence 3: Must be exactly: 'We thank you for your kind consideration of this application.' OUTPUT EXAMPLE: 'For all of the above reasons, we respectfully request that [Name] be granted a [single/multiple-entry] Temporary Resident Permit valid [dates]. [One closing argument sentence]. We thank you for your kind consideration of this application.'",
  "offenceSummaryForHeader": "One-line summary of offence history for header table e.g. 'One (1) DUI conviction (2019, United States)'"
}`;

  try {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': API_KEY,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 3000,
        messages: [{ role: 'user', content: prompt }],
      }),
    });

    if (!response.ok) {
      const err = await response.text();
      return res.status(502).json({ error: `Anthropic error: ${err}` });
    }

    const data = await response.json();
    let raw = data.content[0].text.trim()
      .replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');

    // Extract just the outermost {...} block in case AI adds surrounding text
    const jsonStart = raw.indexOf('{');
    const jsonEnd   = raw.lastIndexOf('}');
    if (jsonStart === -1 || jsonEnd === -1) {
      throw new Error('AI did not return a valid JSON object. Please try generating again.');
    }
    raw = raw.slice(jsonStart, jsonEnd + 1);

    // Fix literal newlines inside JSON string values (common AI mistake)
    // Replace \n that appear inside "..." with \\n
    raw = raw.replace(/"((?:[^"\\]|\\.)*)"/gs, (match, inner) =>
      '"' + inner.replace(/\n/g, '\\n').replace(/\r/g, '') + '"'
    );

    let result;
    try {
      result = JSON.parse(raw);
    } catch (parseErr) {
      console.error('JSON parse failed. Raw AI output:\n', raw.slice(0, 500));
      throw new Error('AI returned malformed JSON. Please try generating again.');
    }
    result.standardOpening = `We represent the above-named individual with respect to ${hisHer} application for a Temporary Resident Permit. An authorization to this effect is enclosed.`;
    if (poeParagraph)   result.poeParagraph   = poeParagraph;
    if (rehabParagraph) result.rehabParagraph = rehabParagraph;
    res.json(result);
  } catch (err) {
    console.error('Generate letter error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─── DOWNLOAD DOCX ───
app.post('/api/download-docx', async (req, res) => {
  const { client, travel, generated, attachments, applicationMode } = req.body;

  // Support both new arrests[] and legacy offences[] formats
  const docxArrests = req.body.arrests || (req.body.offences || []).map(o => ({
    date: o.date || '', country: o.country || '',
    charges: [{ description: o.description || '', outcome: o.outcome || 'conviction',
                sentence: o.sentence || '', foreignStatute: o.foreignStatute || '',
                canadianStatute: o.canadianStatute || '',
                canadianEquivalentLabel: o.canadianEquivalentLabel || '',
                irpa: o.irpa || '' }],
  }));

  function chargeOutcomeLabel(c, withMarker = false) {
    let base;
    if (c.outcome === 'conviction')         return c.sentence || 'Conviction';
    if (c.outcome === 'dropped')            return 'Dropped';
    if (c.outcome === 'provincial')         return c.sentence || 'Provincial Offence';
    const markerMap = { deferred:' *', diversion:' †', dismissed_1203_4:' ‡', cwof_ma:' §', ard_pa:' ¶', pbj_md:' #', expungement:' °' };
    const labelMap  = {
      dismissed: 'Dismissed', dismissed_1203_4: '1203.4 — Dismissed',
      deferred: 'Deferred Adjudication — Dismissed', diversion: 'Diversion Program — Dismissed',
      ard_pa: 'ARD — Dismissed', pbj_md: 'Probation Before Judgment — Dismissed',
      cwof_ma: 'CWOF — Dismissed', expungement: 'Expungement',
    };
    base = labelMap[c.outcome] || c.outcome || '';
    if (withMarker && markerMap[c.outcome]) base += markerMap[c.outcome];
    // Prepend conditions / sentence text if present
    const condText = c.programConditions || (c.outcome === 'dismissed_1203_4' ? c.sentence : '') || '';
    if (condText) base = condText + '\n' + base;
    return base;
  }

  // Build legal notes for deferred/diversion/1203.4 charges (placed after conviction table)
  function buildDocxLegalNotes() {
    let deferredState = null, diversionState = null, cwofFound = false, ardFound = false, pbjFound = false, provincialFound = false;
    const expungementCharges = [];
    const charges1203 = [];
    for (const arr of docxArrests) {
      for (const c of arr.charges || []) {
        if (c.outcome === 'deferred'         && deferredState  === null) deferredState  = arr.state || arr.country || '';
        if (c.outcome === 'diversion'        && diversionState === null) diversionState = arr.state || arr.country || '';
        if (c.outcome === 'cwof_ma')          cwofFound = true;
        if (c.outcome === 'ard_pa')           ardFound  = true;
        if (c.outcome === 'pbj_md')           pbjFound        = true;
        if (c.outcome === 'provincial')       provincialFound = true;
        if (c.outcome === 'expungement')      expungementCharges.push(c);
        if (c.outcome === 'dismissed_1203_4') charges1203.push(c);
      }
    }
    const notes = [];
    const noteStyle = { size: 20 }; // 10pt
    function notePara(text) {
      return new Paragraph({ spacing: { before: 120, after: 80 }, children: [makeText(text, noteStyle)] });
    }
    if (deferredState !== null) {
      const s = deferredState || '[STATE]';
      notes.push(notePara(`* Under ${s} law, upon a plea of guilty and prior to any judgment of guilt, the court may defer proceedings on such conditions as it prescribes, without entering a judgment of guilt. Upon the successful completion of those conditions, the defendant is discharged without a court judgment of guilt, the verdict or plea of guilty is ordered expunged from the record, and the charge is dismissed with prejudice to any further action. It is our position that a deferred sentence under ${s} law is equivalent in nature and effect to a Conditional Discharge under section 730 of the Criminal Code of Canada. IRCC’s policy with respect to conditional discharges is set out in Manual ENF 2/OP 18, Section 14.1, which provides that a conviction does not exist where the court grants an absolute or conditional discharge as provided for under the Criminal Code.`));
      notes.push(notePara(`As our client’s charge was fully dismissed and no finding or judgment of guilt was ever entered, we respectfully submit that this matter has no bearing on our client’s admissibility to Canada, as it does not meet the inadmissibility criteria laid out in Section 36 of IRPA.`));
    }
    if (provincialFound) {
      notes.push(notePara(`We submit that this charge has no equivalent offence under an Act of Parliament in Canada. In the present case, our client's conduct would instead fall under provincial legislation and would not constitute a federal criminal offence. Accordingly, we submit that this incident does not impact our client's admissibility to Canada on the basis of criminality.`));
    }
    if (diversionState !== null) {
      const s = diversionState || '[STATE]';
      notes.push(notePara(`† Under ${s} law, following the successful completion of a pre-trial diversion agreement’s conditions, the charge is subsequently dismissed without a conviction being entered. It is our position that a diversion agreement under ${s} law is equivalent in nature and effect to a Conditional Discharge under section 730 of the Criminal Code of Canada. IRCC’s policy with respect to conditional discharges is set out in Manual ENF 2/OP 18, Section 14.1, which provides that a conviction does not exist where the court grants an absolute or conditional discharge as provided for under the Criminal Code.`));
      notes.push(notePara(`As our client’s charge was fully dismissed and no finding or judgment of guilt was ever entered, we respectfully submit that this matter has no bearing on our client’s admissibility to Canada, as it does not meet the inadmissibility criteria laid out in Section 36 of IRPA.`));
    }
    if (cwofFound) {
      const nameParts = (client.name || '').trim().split(/\s+/);
      const lastName  = nameParts.length > 1 ? nameParts[nameParts.length - 1] : (nameParts[0] || '[CLIENT]');
      const titleStr  = (client.gender === 'ms') ? 'Ms.' : 'Mr.';
      const titleLast = `${titleStr} ${lastName}`;
      notes.push(notePara(`§ We respectfully contend that a "Continued Without a Finding" disposition in the state of Massachusetts is not a criminal conviction and is equivalent to a Conditional Discharge under Section 730 of the Criminal Code of Canada. IRCC's Policy with respect to conditional discharge is set forth in Manual ENF 2/OP 18, Section 14.2: A conviction does not exist [when] the court grants an absolute or conditional discharge as provided for in the Criminal Code.`));
      notes.push(notePara(`Additionally, we contend that ${titleLast} would not be inadmissible under the "committing of an act" provision. According to ENF 2/OP 18, Section 3.9, the "committing an act" provision should not be used when:`));
      notes.push(notePara(`—  in most cases, when authorities in the foreign jurisdiction indicate they would not lay a charge or make known to an officer their decision or intent to drop the charges;`));
      notes.push(notePara(`—  the trial is concluded and no conviction results (for example, acquittal, discharge, deferral);`));
    }
    if (pbjFound) {
      notes.push(notePara(`# Under Maryland law, Probation Before Judgment (PBJ) is a special legal disposition that allows a judge to strike a guilty plea or verdict and instead place a defendant on probation without entering a formal conviction. Upon successful completion of probation, the case is closed without a conviction being entered. It is our position that a PBJ is equivalent in nature and effect to a Conditional Discharge under section 730 of the Criminal Code of Canada. IRCC's policy with respect to conditional discharges is set out in Manual ENF 2/OP 18, Section 14.1, which provides that a conviction does not exist where the court grants an absolute or conditional discharge as provided for under the Criminal Code.`));
      notes.push(notePara(`As our client's charge was fully dismissed and no finding or judgment of guilt was ever entered, we respectfully submit that this matter has no bearing on our client's admissibility to Canada, as it does not meet the inadmissibility criteria laid out in Section 36 of IRPA.`));
    }
    if (ardFound) {
      const nameParts = (client.name || '').trim().split(/\s+/);
      const lastName  = nameParts.length > 1 ? nameParts[nameParts.length - 1] : (nameParts[0] || '[CLIENT]');
      const titleStr  = (client.gender === 'ms') ? 'Ms.' : 'Mr.';
      const heShe     = (client.gender === 'ms') ? 'she' : 'he';
      const hisHer    = (client.gender === 'ms') ? 'her' : 'his';
      const titleLast = `${titleStr} ${lastName}`;
      notes.push(notePara(`¶ ${titleLast} was admitted into the pre-trial intervention program Accelerated Rehabilitative Disposition (ARD). This program is exclusively offered to individuals who have no prior criminal history.`));
      notes.push(notePara(`Given that ${heShe} was accepted into the ARD Program, upon completion of all the terms of ${hisHer} sentence, ${titleLast}'s aforementioned charge became eligible for expungement, per enclosed document supporting the completion of an ARD Program.`));
      notes.push(notePara(`As indicated in the enclosed documentation, conditional to the completion of the ARD, the case is dismissed without a finding of guilt, with a possibility of expungement available. Therefore, since ${titleLast} completed the terms of ${hisHer} ARD program, this singular charge does not indicate criminality and does not inhibit ${hisHer} admissibility into Canada.`));
    }
    for (const c of expungementCharges) {
      const dateStr = c.expungementDate
        ? new Date(c.expungementDate + 'T12:00:00').toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
        : '[EXPUNGEMENT DATE]';
      const statute = c.expungementStatute || '[STATUTE]';
      notes.push(notePara(`° On ${dateStr}, our client obtained an expungement of this conviction pursuant to ${statute}. As a result of that order, the matter is deemed never to have occurred, and our client and all criminal justice agencies are entitled to reply to any inquiry regarding the matter by stating that no such action ever took place.`));
      notes.push(notePara(`IRCC's policy with respect to expungements is set out in Manual ENF 2/OP 18, Section 14.2, which provides that an expungement does not constitute a conviction, defining the term as meaning to strike out, obliterate, mark for deletion, efface completely, or deem never to have occurred.`));
      notes.push(notePara(`Accordingly, we respectfully submit that, as the conviction has been expunged, it does not constitute an incident to be considered under section 36 of IRPA for the purposes of assessing our client's admissibility to Canada.`));
    }
    for (const c of charges1203) {
      const dateStr = c.expungementDate
        ? new Date(c.expungementDate).toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
        : '[EXPUNGEMENT DATE]';
      const clientFullName = client.name || '[CLIENT NAME]';
      notes.push(notePara(`‡ On ${dateStr}, the Superior Court of California granted relief and dismissed this conviction under Section 1203.4 of the California Penal Code. This order resulted in the withdrawal of ${clientFullName}’s plea of "guilty" and a "not guilty" plea being entered in its place. It is the contention of our office that a dismissal pursuant to Section 1203.4 of the California Penal Code is the equivalent of a record suspension or an expungement under Canadian law due to its similar content, aim, and effect. It follows that, according to IRCC regulations on determining admissibility, specifically ENF/OP18: Evaluating Inadmissibility, an expunged conviction is to be treated as "Not a conviction. Expunged means to strike out; obliterate; mark for deletion; to efface completely; deemed to have never occurred."`));
      notes.push(notePara(`In Canada, section 36(3)(b) of the Immigration and Refugee Protection Act states that criminal inadmissibility "may not be based on a conviction in respect of which a record suspension has been ordered…". For these reasons, it is the contention of our office that a dismissal under Section 1203.4, equating to a record suspension or expungement, does not affect our client’s admissibility to Canada based on criminality.`));
    }
    return notes;
  }

  const CM_INITIALS = { Dimitra: 'dm', Sarah: 'sk', Shealeigh: 'ss', Sophie: 'sp', Diane: 'dim' };
  const cmInitials = CM_INITIALS[client.caseManager] || (client.caseManager ? client.caseManager.toLowerCase().slice(0,2) : '');
  const initialsLine = cmInitials ? `DL/${cmInitials}` : 'DL';

  function formatReDocx(name, dob) {
    const parts = (name || '').trim().split(/\s+/);
    const last  = parts.pop().toUpperCase();
    const first = parts.join(' ') || last;
    let dobStr = '';
    if (dob) {
      const d = new Date(dob + 'T12:00:00');
      dobStr = ` (DOB: ${d.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })})`;
    }
    return `${last}, ${first}${dobStr}`;
  }

  const todayStr = new Date().toLocaleDateString('en-US', {
    year: 'numeric', month: 'long', day: 'numeric'
  });
  const isSerious = docxArrests.some(arr => (arr.charges || []).some(c => c.irpa && /36\(1\)/.test(c.irpa)));
  const criminalityLabel = isSerious ? 'Serious' : 'Non-serious';

  function makeText(text, opts = {}) {
    return new TextRun({
      text,
      font: 'Times New Roman',
      size: 24, // 12pt
      ...opts,
    });
  }

  // Splits text on \n and returns an array of TextRun children with line breaks
  function makeMultilineRuns(text, opts = {}) {
    const lines = (text || '').split('\n');
    const runs = [];
    lines.forEach((line, i) => {
      if (i > 0) runs.push(new TextRun({ break: 1 }));
      runs.push(makeText(line, opts));
    });
    return runs;
  }

  // Like fmtMultiDocx but returns an array of TextRun children preserving \n line breaks
  function fmtMultiDocxRuns(values, opts = {}) {
    const nonEmpty = values.map(v => v || '—');
    const lines = nonEmpty.length === 1
      ? nonEmpty[0]
      : nonEmpty.map((v, idx) => `${idx + 1}. ${v}`).join('\n');
    return makeMultilineRuns(lines, opts);
  }

  function para(text, opts = {}) {
    return new Paragraph({
      alignment: opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
      spacing: { after: 160, line: 480, lineRule: 'auto' }, // 2.0 line spacing
      children: [makeText(text, opts)],
    });
  }

  function sectionHeader(text) {
    return new Paragraph({
      spacing: { before: 300, after: 100 },
      children: [makeText(text, { bold: true, underline: { type: UnderlineType.SINGLE } })],
    });
  }

  function emptyLine() {
    return new Paragraph({ children: [makeText('')], spacing: { after: 100 } });
  }

  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

  // ── Header Table (borderless, tab-aligned) ──
  function headerCell(label, value) {
    return new TableRow({
      children: [
        new TableCell({
          width: { size: 20, type: WidthType.PERCENTAGE },
          borders: noBorders,
          children: [new Paragraph({ spacing: { after: 60 }, children: [makeText(label, { bold: true })] })],
        }),
        new TableCell({
          width: { size: 80, type: WidthType.PERCENTAGE },
          borders: noBorders,
          children: [new Paragraph({ spacing: { after: 60 }, children: [makeText(value || '')] })],
        }),
      ],
    });
  }

  const headerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      headerCell('Attention:', 'Canadian Immigration Officer'),
      headerCell('RE:', formatReDocx(client.name, client.dob)),
      headerCell('Subject:', 'Application for a Temporary Resident Permit'),
      headerCell('Purpose:', travel.purpose || ''),
      headerCell('Entry Type:', travel.entryType || ''),
      headerCell('Duration:', travel.duration || ''),
      headerCell('Criminality:', criminalityLabel),
    ],
  });

  // ── Conviction Table ──
  function fmtMultiDocx(values) {
    const nonEmpty = values.map(v => v || '—');
    if (nonEmpty.length === 1) return nonEmpty[0];
    return nonEmpty.map((v, idx) => `${idx+1}. ${v}`).join('\n');
  }

  const convictionRows = [
    new TableRow({
      tableHeader: true,
      children: ['Date', 'Offence', 'Sentence', 'Foreign Statute', 'Canadian Equivalent'].map(h =>
        new TableCell({
          shading: { fill: 'D9D9D9' },
          children: [new Paragraph({ children: [makeText(h, { bold: true })] })],
        })
      ),
    }),
    ...docxArrests.map(arr => {
      const descs    = (arr.charges || []).map(c => c.description);
      const sents    = (arr.charges || []).map(c => chargeOutcomeLabel(c, true));
      const statutes = (arr.charges || []).map(c => c.foreignStatute);
      const equivs   = (arr.charges || []).map(c => (c.outcome === 'dropped' || c.outcome === 'provincial') ? 'N/A' : (c.canadianEquivalentLabel || c.canadianStatute));
      return new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [makeText(arr.date || '')] })] }),
          new TableCell({ children: [new Paragraph({ children: [makeText(fmtMultiDocx(descs))] })] }),
          new TableCell({ children: [new Paragraph({ children: fmtMultiDocxRuns(sents) })] }),
          new TableCell({ children: [new Paragraph({ children: [makeText(fmtMultiDocx(statutes))] })] }),
          new TableCell({ children: [new Paragraph({ children: [makeText(fmtMultiDocx(equivs))] })] }),
        ],
      });
    }),
  ];

  const convictionTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: convictionRows,
  });

  // ── No-records notes (italic, after table) — grouped by reason ──
  function docxDateList(arrs) {
    const dates = arrs.map(a => a.date ? `(${a.date})` : '');
    if (dates.length === 1) return `pertaining to the ${dates[0]} arrest`;
    if (dates.length === 2) return `pertaining to the ${dates[0]} and ${dates[1]} arrests`;
    return `pertaining to the ${dates.slice(0,-1).join(', ')}, and ${dates[dates.length-1]} arrests`;
  }
  const docxPurged   = docxArrests.filter(a => a.noRecords && a.noRecordsReason !== 'untimely');
  const docxUntimely = docxArrests.filter(a => a.noRecords && a.noRecordsReason === 'untimely');
  const noRecordsTexts = [];
  if (docxPurged.length)
    noRecordsTexts.push(`Please note that court records ${docxDateList(docxPurged)} are not available, as they have been purged.`);
  if (docxUntimely.length)
    noRecordsTexts.push(`Please note that court records ${docxDateList(docxUntimely)} are not available, as the client was unable to obtain them in time for this submission.`);
  const noRecordsParagraphs = noRecordsTexts.map(text => new Paragraph({
        spacing: { before: 80, after: 80 },
        children: [new TextRun({ text, font: 'Times New Roman', size: 20, italics: true })],
      }));

  // ── Attachments ──
  const attachmentList = (attachments || []).map(a =>
    new Paragraph({
      bullet: { level: 0 },
      children: [makeText(a)],
    })
  );

  // ── Footnotes ──
  const footnotes = (generated.economicBenefitFootnotes || []).map((fn, i) =>
    new Paragraph({
      spacing: { after: 100 },
      children: [makeText(`[${i + 1}] ${fn}`, { size: 20, italics: true })],
    })
  );

  // ─── BODY TEXT PARAGRAPHS ───
  function bodyParas(text) {
    if (!text) return [emptyLine()];
    return text.split('\n').filter(t => t.trim()).map(t => para(t));
  }

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: { top: 1440, right: 1080, bottom: 1440, left: 1080 },
        },
      },
      children: [
        // Letterhead
        // ── 3-column letterhead ──
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                // Left: address
                new TableCell({
                  width: { size: 33, type: WidthType.PERCENTAGE },
                  borders: noBorders,
                  verticalAlign: VerticalAlign.CENTER,
                  children: [
                    new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: '420 NOTRE DAME W, SUITE 310', font: 'Times New Roman', size: 18 })] }),
                    new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: 'MONTREAL, QUEBEC', font: 'Times New Roman', size: 18 })] }),
                    new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: 'CANADA, H2Y 1V3', font: 'Times New Roman', size: 18 })] }),
                  ],
                }),
                // Center: branding — leaf | vertical bar | firm name
                new TableCell({
                  width: { size: 34, type: WidthType.PERCENTAGE },
                  borders: noBorders,
                  verticalAlign: VerticalAlign.CENTER,
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { after: 0 },
                      children: [
                        new TextRun({ text: '🍁', font: 'Times New Roman', size: 22, color: '333333' }),
                        new TextRun({ text: '  |  ', font: 'Times New Roman', size: 22, color: '555555' }),
                        new TextRun({ text: 'COHEN', font: 'Times New Roman', size: 22, bold: true, color: '222222' }),
                        new TextRun({ text: ' IMMIGRATION LAW', font: 'Times New Roman', size: 22, color: '333333' }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { after: 0 },
                      children: [new TextRun({ text: 'S I N C E  1 9 7 6', font: 'Times New Roman', size: 14, color: '777777' })],
                    }),
                  ],
                }),
                // Right: contact
                new TableCell({
                  width: { size: 33, type: WidthType.PERCENTAGE },
                  borders: noBorders,
                  verticalAlign: VerticalAlign.CENTER,
                  children: [
                    new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0 }, children: [new TextRun({ text: 'TEL: (514) 937-9445', font: 'Times New Roman', size: 18 })] }),
                    new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0 }, children: [new TextRun({ text: 'FAX: (514) 937-2618', font: 'Times New Roman', size: 18 })] }),
                    new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0 }, children: [new TextRun({ text: 'cohenlaw@canadavisa.com', font: 'Times New Roman', size: 18 })] }),
                  ],
                }),
              ],
            }),
          ],
        }),
        // Date + Without Prejudice
        new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          spacing: { after: 0 },
          children: [
            makeText(todayStr),
            new TextRun({ text: '\t', font: 'Times New Roman', size: 24 }),
            makeText('Without Prejudice', { italics: true }),
          ],
        }),

        // Montreal note
        new Paragraph({
          spacing: { after: 240 },
          children: [makeText('Montreal')],
        }),

        emptyLine(),

        // Header table
        headerTable,

        emptyLine(),
        emptyLine(),

        // Salutation
        para('Dear Sir/Madam,'),

        emptyLine(),

        // Introduction
        ...bodyParas(generated.introduction),

        emptyLine(),

        // Section 1
        sectionHeader('1. Foreign Offence(s)'),
        emptyLine(),
        convictionTable,
        ...buildDocxLegalNotes(),
        ...noRecordsParagraphs,

        emptyLine(),

        // Section 2
        sectionHeader('2. Reason for Request'),
        ...bodyParas(generated.reasonForRequest),

        emptyLine(),

        // Section 3
        sectionHeader('3. Economic Benefit to Canada'),
        ...bodyParas(generated.economicBenefit),

        emptyLine(),

        // Section 4
        sectionHeader('4. Rehabilitation Factors'),
        ...bodyParas(generated.rehabilitationFactors),

        emptyLine(),

        // Section 5
        sectionHeader('5. Conclusion'),
        ...bodyParas(generated.conclusion),

        emptyLine(),
        emptyLine(),

        // Attachments
        sectionHeader('Attachments:'),
        new Paragraph({ spacing: { after: 100 }, children: [makeText('A. Supporting Documents:')] }),
        ...(attachmentList.length ? attachmentList : [para('• Personal statement of the applicant')]),

        emptyLine(),
        emptyLine(),

        // Sign-off
        para('Respectfully yours,'),
        emptyLine(),
        para('COHEN IMMIGRATION LAW', { bold: true }),
        new Paragraph({
          spacing: { after: 60 },
          children: [new ImageRun({
            data: fs.readFileSync(path.join(__dirname, 'signature.png')),
            transformation: { width: 160, height: 72 },
            type: 'png',
          })],
        }),
        para('Per: Attorney Daniel Levy'),
        para(initialsLine),

        emptyLine(),
        emptyLine(),

        // Footnotes
        ...(footnotes.length ? [
          new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 1 } },
            spacing: { before: 200, after: 100 },
            children: [makeText('Sources:', { bold: true, size: 20 })],
          }),
          ...footnotes,
        ] : []),
      ],
    }],
  });

  try {
    const buffer = await Packer.toBuffer(doc);
    const filename = `TRP_${(client.name || 'Client').replace(/\s+/g, '_')}_${new Date().toISOString().slice(0,10)}.docx`;
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(buffer);
  } catch (err) {
    console.error('DOCX error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
//  LETTERS — Supabase persistence
// ─────────────────────────────────────────────

app.get('/api/letters', async (req, res) => {
  const { data, error } = await supabase
    .from('letters')
    .select('*')
    .order('saved_at', { ascending: false });
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.post('/api/letters', async (req, res) => {
  const b = req.body;
  const row = {
    id:              b.id,
    saved_at:        b.savedAt,
    case_manager:    b.caseManager,
    client_name:     b.clientName,
    client_dob:      b.clientDob,
    offence_summary: b.offenceSummary,
    travel_purpose:  b.travelPurpose,
    entry_type:      b.entryType,
    date_from:       b.dateFrom,
    date_to:         b.dateTo,
    payload:         b.payload,
    generated:       b.generated,
  };
  const { data, error } = await supabase
    .from('letters')
    .upsert(row, { onConflict: 'id' })
    .select()
    .single();
  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.delete('/api/letters/:id', async (req, res) => {
  const { error } = await supabase
    .from('letters')
    .delete()
    .eq('id', req.params.id);
  if (error) return res.status(500).json({ error: error.message });
  res.json({ ok: true });
});

app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

if (require.main === module) {
  app.listen(PORT, () => console.log(`TRP Generator running at http://localhost:${PORT}`));
}

module.exports = app;
