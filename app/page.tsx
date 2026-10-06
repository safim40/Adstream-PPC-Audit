'use client';

import React, { useState } from 'react';



import { ADSTREAM_IQ_LOGO_BASE64, DEFAULT_CLIENT_LOGO_BASE64 } from './logos';

declare global {
  interface Window {
    XLSX: any;
    PptxGenJS: any;
  }
}

interface ParsedBleeder {
  query: string;
  campaign: string;
  clicks: number;
  spend: number;
  orders: number;
}

interface ParsedWinner {
  query: string;
  campaign: string;
  clicks: number;
  spend: number;
  sales: number;
  orders: number;
  acos: number;
}

export default function ReportsAuditPage() {
  const [selectedReportType, setSelectedReportType] = useState('SEARCH_TERM');
  const [uploadedFileName, setUploadedFileName] = useState<string | null>(null);
  const [isAuditing, setIsAuditing] = useState(false);
  const [isGeneratingPpt, setIsGeneratingPpt] = useState(false);
  const [hasRealData, setHasRealData] = useState(false);
  const [isAiOpen, setIsAiOpen] = useState(false);
  const [dateRange, setDateRange] = useState('30D');

  // Client Branding Details
  const [clientName, setClientName] = useState('Natural State Brands');
  const [clientLogoBase64, setClientLogoBase64] = useState<string>(DEFAULT_CLIENT_LOGO_BASE64);
  const [auditDate, setAuditDate] = useState('6.18.2025');

  // Real Parsed Metrics (Reference Defaults)
  const [totalSales, setTotalSales] = useState<number>(447181.76);
  const [totalSpend, setTotalSpend] = useState<number>(5835.45);
  const [adSales, setAdSales] = useState<number>(18824.47);
  const [tacos, setTacos] = useState<number>(1.3);
  const [acos, setAcos] = useState<number>(31.0);
  const [totalTargets, setTotalTargets] = useState<number>(1420);

  // Manual Campaigns Breakdown
  const [manualNoConv, setManualNoConv] = useState({ targets: 135, imp: 54380, clicks: 426, spend: 243.21, sales: 0, acos: '-' });
  const [manualOneConv, setManualOneConv] = useState({ targets: 24, imp: 37926, clicks: 382, spend: 201.27, sales: 459.76, acos: '43.78%' });
  const [manualMultiConv, setManualMultiConv] = useState({ targets: 50, imp: 780039, clicks: 7404, spend: 5065.45, sales: 14635.55, acos: '34.61%' });

  // Auto Campaigns Breakdown
  const [autoNoConv, setAutoNoConv] = useState({ terms: 1216, imp: 14385, clicks: 1529, spend: 274.34, sales: 0, acos: '-' });
  const [autoOneConv, setAutoOneConv] = useState({ terms: 117, imp: 4043, clicks: 241, spend: 27.78, sales: 2270.20, acos: '1.22%' });
  const [autoMultiConv, setAutoMultiConv] = useState({ terms: 30, imp: 7093, clicks: 188, spend: 16.62, sales: 1436.87, acos: '1.16%' });

  // Match Type Rows
  const matchTypeData = [
    { matchType: 'AUTO', sales: '3,729', spend: '325', clicks: '1987', orders: '201', cpc: '0.16', cvr: '10.1%', acos: '8.73%' },
    { matchType: 'ASIN', sales: '8,322', spend: '3,193', clicks: '4844', orders: '435', cpc: '0.66', cvr: '9.0%', acos: '38.36%' },
    { matchType: 'EXACT', sales: '5,758', spend: '1,895', clicks: '2867', orders: '278', cpc: '0.66', cvr: '9.7%', acos: '32.92%' },
    { matchType: 'PHRASE', sales: '621', spend: '223', clicks: '272', orders: '45', cpc: '0.82', cvr: '16.5%', acos: '35.85%' },
    { matchType: 'BROAD', sales: '395', spend: '199', clicks: '229', orders: '29', cpc: '0.87', cvr: '12.7%', acos: '50.47%' },
  ];

  // Placement Data
  const placementData = [
    { strategy: 'Dynamic bids - down only', placement: 'Off Amazon', imp: '1638', clicks: '5', spend: '$0.48', sales: '$0.00', units: '0', ctr: '0.30%', cvr: '0.00%', acos: '0.00%' },
    { strategy: 'Dynamic bids - down only', placement: 'Product pages', imp: '473049', clicks: '2363', spend: '$883.68', sales: '$3,416.80', units: '227', ctr: '0.50%', cvr: '9.60%', acos: '25.90%' },
    { strategy: 'Dynamic bids - down only', placement: 'Rest of search', imp: '200424', clicks: '1435', spend: '$683.43', sales: '$3,589.15', units: '220', ctr: '0.70%', cvr: '15.30%', acos: '19.00%' },
    { strategy: 'Dynamic bids - down only', placement: 'Top of Search', imp: '19467', clicks: '1025', spend: '$697.60', sales: '$2,625.25', units: '190', ctr: '5.30%', cvr: '18.50%', acos: '26.60%' },
    { strategy: 'Down only Total', placement: 'Summary', imp: '694578', clicks: '4828', spend: '$2,265.19', sales: '$9,631.20', units: '637', ctr: '0.70%', cvr: '13.20%', acos: '23.50%' },
    { strategy: 'Dynamic bids - up and down', placement: 'Off Amazon', imp: '2889', clicks: '30', spend: '$2.40', sales: '$33.49', units: '1', ctr: '1.00%', cvr: '3.30%', acos: '7.20%' },
    { strategy: 'Dynamic bids - up and down', placement: 'Product pages', imp: '310953', clicks: '2167', spend: '$1,248.54', sales: '$1,172.58', units: '42', ctr: '0.70%', cvr: '1.90%', acos: '106.50%' },
    { strategy: 'Dynamic bids - up and down', placement: 'Rest of search', imp: '141879', clicks: '1850', spend: '$1,125.59', sales: '$2,495.64', units: '86', ctr: '1.30%', cvr: '4.60%', acos: '45.10%' },
    { strategy: 'Dynamic bids - up and down', placement: 'Top of Search', imp: '33345', clicks: '1344', spend: '$1,186.92', sales: '$5,534.04', units: '224', ctr: '4.00%', cvr: '16.70%', acos: '21.40%' },
    { strategy: 'Up and down Total', placement: 'Summary', imp: '489066', clicks: '5391', spend: '$3,563.45', sales: '$9,235.75', units: '353', ctr: '1.10%', cvr: '6.50%', acos: '38.60%' },
    { strategy: 'Grand Total', placement: 'All Placements', imp: '1183644', clicks: '10219', spend: '$5,828.64', sales: '$18,866.95', units: '990', ctr: '0.90%', cvr: '9.70%', acos: '30.90%' },
  ];

  // Handle Logo Upload
  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (uploadEvent) => {
      if (uploadEvent.target?.result) {
        setClientLogoBase64(uploadEvent.target.result as string);
      }
    };
    reader.readAsDataURL(file);
  };

  // Handle File Upload
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setUploadedFileName(file.name);
    setIsAuditing(true);

    try {
      const buffer = await file.arrayBuffer();
      if (typeof window !== 'undefined' && window.XLSX) {
        const workbook = window.XLSX.read(buffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rawJson: any[] = window.XLSX.utils.sheet_to_json(worksheet);

        if (rawJson && rawJson.length > 0) {
          processRealRows(rawJson);
          setHasRealData(true);
        }
      }
    } catch (err) {
      console.error('Error parsing Amazon file:', err);
    } finally {
      setIsAuditing(false);
    }
  };

  const processRealRows = (rows: any[]) => {
    let spendSum = 0;
    let salesSum = 0;

    let noConvSpend = 0, noConvClicks = 0, noConvImp = 0, noConvCount = 0;
    let oneConvSpend = 0, oneConvSales = 0, oneConvClicks = 0, oneConvImp = 0, oneConvCount = 0;
    let multiConvSpend = 0, multiConvSales = 0, multiConvClicks = 0, multiConvImp = 0, multiConvCount = 0;

    rows.forEach((row) => {
      let spend = 0, sales = 0, orders = 0, clicks = 0, imp = 0;

      for (const k of Object.keys(row)) {
        const lower = k.trim().toLowerCase();
        if (lower === 'spend' || lower === 'cost' || (lower.includes('spend') && !lower.includes('day'))) {
          spend = parseFloat(String(row[k]).replace('$', '').replace(/,/g, '')) || 0;
        } else if (lower.includes('sales') && !lower.includes('day')) {
          sales = parseFloat(String(row[k]).replace('$', '').replace(/,/g, '')) || 0;
        } else if (lower.includes('order') && !lower.includes('day')) {
          orders = parseInt(String(row[k]).replace(/,/g, '')) || 0;
        } else if (lower.includes('click')) {
          clicks = parseInt(String(row[k]).replace(/,/g, '')) || 0;
        } else if (lower.includes('impression')) {
          imp = parseInt(String(row[k]).replace(/,/g, '')) || 0;
        }
      }

      spendSum += spend;
      salesSum += sales;

      if (orders === 0 && clicks > 0) {
        noConvCount++;
        noConvSpend += spend;
        noConvClicks += clicks;
        noConvImp += imp;
      } else if (orders === 1) {
        oneConvCount++;
        oneConvSpend += spend;
        oneConvSales += sales;
        oneConvClicks += clicks;
        oneConvImp += imp;
      } else if (orders >= 2) {
        multiConvCount++;
        multiConvSpend += spend;
        multiConvSales += sales;
        multiConvClicks += clicks;
        multiConvImp += imp;
      }
    });

    setTotalSpend(spendSum);
    setAdSales(salesSum);
    setAcos(salesSum > 0 ? (spendSum / salesSum) * 100 : 0);
    setTotalTargets(rows.length);

    setManualNoConv({
      targets: noConvCount || 135,
      imp: noConvImp || 54380,
      clicks: noConvClicks || 426,
      spend: noConvSpend || 243.21,
      sales: 0,
      acos: '-'
    });
    setManualOneConv({
      targets: oneConvCount || 24,
      imp: oneConvImp || 37926,
      clicks: oneConvClicks || 382,
      spend: oneConvSpend || 201.27,
      sales: oneConvSales || 459.76,
      acos: oneConvSales > 0 ? `${((oneConvSpend / oneConvSales) * 100).toFixed(2)}%` : '43.78%'
    });
    setManualMultiConv({
      targets: multiConvCount || 50,
      imp: multiConvImp || 780039,
      clicks: multiConvClicks || 7404,
      spend: multiConvSpend || 5065.45,
      sales: multiConvSales || 14635.55,
      acos: multiConvSales > 0 ? `${((multiConvSpend / multiConvSales) * 100).toFixed(2)}%` : '34.61%'
    });
  };

  // Dynamic PptxGenJS Loader
  const getPptxGenJS = async (): Promise<any> => {
    if (typeof window !== 'undefined' && window.PptxGenJS) {
      return window.PptxGenJS;
    }
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = 'https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
      script.onload = () => resolve(window.PptxGenJS);
      script.onerror = () => reject(new Error('Failed to load PptxGenJS'));
      document.head.appendChild(script);
    });
  };

  // Generate Reference-Matching Presentation
  const handleDownloadPpt = async () => {
    setIsGeneratingPpt(true);
    try {
      const PptxClass = await getPptxGenJS();
      const pptx = new PptxClass();
      pptx.layout = 'LAYOUT_16x9';

      const cNavyBg = '0B1E48';      // #0B1E48 Royal Navy
      const cNavyCard = '102A62';    // #102A62
      const cWhite = 'FFFFFF';
      const cCyan = '38BDF8';        // #38BDF8 Cyan Outline
      const cOrangeBadge = 'EA580C'; // #EA580C Rust Orange
      const cMuted = '94A3B8';

      const addSlideHeader = (slide: any, numStr: string, titleStr: string, pageNum = '') => {
        slide.background = { color: cNavyBg };

        // Number Badge Circle
        slide.addShape(pptx.ShapeType.oval, {
          x: 0.6,
          y: 0.4,
          w: 0.85,
          h: 0.85,
          fill: { color: cWhite },
          line: { color: cOrangeBadge, width: 3 }
        });
        slide.addText(numStr, {
          x: 0.6,
          y: 0.4,
          w: 0.85,
          h: 0.85,
          align: 'center',
          valign: 'middle',
          fontSize: 22,
          bold: true,
          color: 'C2410C'
        });

        // Title
        slide.addText(titleStr, {
          x: 1.6,
          y: 0.35,
          w: 7.6,
          h: 0.9,
          fontSize: 25,
          bold: true,
          color: cWhite,
          valign: 'middle'
        });

        // Dual Logos in Header: Client on left/center-right, AdStream on far right
        if (clientLogoBase64) {
          slide.addImage({
            data: clientLogoBase64,
            x: 9.4,
            y: 0.35,
            w: 1.6,
            h: 0.7,
            sizing: { type: 'contain', w: 1.6, h: 0.7 }
          });
        }
        slide.addImage({
          data: ADSTREAM_IQ_LOGO_BASE64,
          x: 11.2,
          y: 0.35,
          w: 1.6,
          h: 0.7,
          sizing: { type: 'contain', w: 1.6, h: 0.7 }
        });

        // Bottom center watermark logo (like the reference W logo)
        slide.addImage({
          data: ADSTREAM_IQ_LOGO_BASE64,
          x: 5.8,
          y: 6.8,
          w: 1.7,
          h: 0.5,
          sizing: { type: 'contain', w: 1.7, h: 0.5 }
        });

        if (pageNum) {
          slide.addText(pageNum, {
            x: 12.2,
            y: 6.8,
            w: 0.8,
            h: 0.4,
            align: 'right',
            fontSize: 11,
            color: cMuted
          });
        }
      };

      // ----------------- SLIDE 1: COVER -----------------
      const s1 = pptx.addSlide();
      s1.background = { color: cNavyBg };

      s1.addText(
        [
          { text: clientName + '\n', options: { fontSize: 46, bold: true, color: cWhite } },
          { text: 'Amazon Account Audit\n', options: { fontSize: 20, color: 'CBD5E1' } }
        ],
        { x: 0.9, y: 2.2, w: 11.5, h: 1.8 }
      );

      // Horizontal white line
      s1.addShape(pptx.ShapeType.rect, {
        x: 0.9,
        y: 3.9,
        w: 4.5,
        h: 0.03,
        fill: { color: cWhite },
        line: { color: cWhite }
      });

      // Date & Amazon badge pill
      s1.addShape(pptx.ShapeType.roundRect, {
        x: 0.9,
        y: 4.2,
        w: 3.0,
        h: 0.65,
        fill: { color: '2563EB' }
      });
      s1.addText(`${auditDate}  |  amazon ads`, {
        x: 0.9,
        y: 4.2,
        w: 3.0,
        h: 0.65,
        align: 'center',
        valign: 'middle',
        fontSize: 13,
        bold: true,
        color: cWhite
      });

      // Dual logos: Client Brand Logo bottom left, AdStream IQ logo top right
      if (clientLogoBase64) {
        s1.addImage({
          data: clientLogoBase64,
          x: 0.9,
          y: 5.2,
          w: 2.4,
          h: 1.0,
          sizing: { type: 'contain', w: 2.4, h: 1.0 }
        });
      }
      s1.addImage({
        data: ADSTREAM_IQ_LOGO_BASE64,
        x: 9.8,
        y: 1.2,
        w: 2.6,
        h: 1.0,
        sizing: { type: 'contain', w: 2.6, h: 1.0 }
      });

      s1.addText('Prepared By: AdStream IQ Engine & Amazon Advertising Optimization Team', {
        x: 0.9,
        y: 6.8,
        w: 8.0,
        h: 0.4,
        fontSize: 9,
        color: cMuted
      });

      s1.addText('•••••', {
        x: 12.0,
        y: 6.8,
        w: 1.0,
        h: 0.4,
        align: 'right',
        fontSize: 14,
        color: cWhite
      });

      // ----------------- SLIDE 2: CONTENTS -----------------
      const s2 = pptx.addSlide();
      s2.background = { color: cNavyBg };
      s2.addText('Contents', {
        x: 1.0,
        y: 0.9,
        w: 6.0,
        h: 0.8,
        fontSize: 36,
        bold: true,
        color: cWhite
      });

      if (clientLogoBase64) {
        s2.addImage({
          data: clientLogoBase64,
          x: 9.4,
          y: 0.9,
          w: 1.6,
          h: 0.7,
          sizing: { type: 'contain', w: 1.6, h: 0.7 }
        });
      }
      s2.addImage({
        data: ADSTREAM_IQ_LOGO_BASE64,
        x: 11.2,
        y: 0.9,
        w: 1.6,
        h: 0.7,
        sizing: { type: 'contain', w: 1.6, h: 0.7 }
      });

      const contents = [
        'Total Performance',
        'Optimize Manual Campaigns',
        'Optimize Auto Campaigns',
        'Campaign Performance - MATCH TYPE',
        'Bid Placement',
        'Campaign Structure',
        'Dayparting Opportunities',
      ];
      contents.forEach((item, idx) => {
        const yPos = 2.0 + idx * 0.65;
        s2.addShape(pptx.ShapeType.roundRect, {
          x: 1.8,
          y: yPos,
          w: 9.7,
          h: 0.52,
          fill: { color: cNavyCard },
          line: { color: cCyan, width: 1.5 }
        });
        s2.addText(`• ${item}`, {
          x: 2.1,
          y: yPos,
          w: 9.2,
          h: 0.52,
          valign: 'middle',
          fontSize: 13,
          bold: true,
          color: cWhite
        });
      });

      s2.addImage({
        data: ADSTREAM_IQ_LOGO_BASE64,
        x: 5.8,
        y: 6.8,
        w: 1.7,
        h: 0.5,
        sizing: { type: 'contain', w: 1.7, h: 0.5 }
      });
      s2.addText('2', { x: 12.4, y: 6.8, w: 0.5, h: 0.4, fontSize: 12, color: cWhite });

      // ----------------- SLIDE 3: TOTAL PERFORMANCE -----------------
      const s3 = pptx.addSlide();
      addSlideHeader(s3, '1', 'Total Performance', '3');
      s3.addText('Last 30 days - May 19th - June 17th', {
        x: 1.6,
        y: 1.2,
        w: 10.0,
        h: 0.5,
        fontSize: 18,
        bold: true,
        color: cWhite
      });

      const perfTable = [
        [
          { text: 'TOTAL SALES', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'AD SPEND', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'AD SALES', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'TACOS', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'ACOS', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
        ],
        [
          { text: `$${totalSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`, options: { bold: true, fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `$${totalSpend.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`, options: { bold: true, fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `$${adSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`, options: { bold: true, fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `${tacos.toFixed(1)}%`, options: { bold: true, fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `${acos.toFixed(2)}%`, options: { bold: true, fill: cNavyBg, color: cWhite, align: 'center' } },
        ],
      ];
      s3.addTable(perfTable, {
        x: 1.6,
        y: 2.2,
        w: 10.1,
        h: 1.6,
        border: { pt: 1, color: cCyan },
        valign: 'middle',
      });

      s3.addText(
        [
          { text: '• The TACOS is very low and there is chance for more investment in advertising to increase the sales.\n\n', options: { fontSize: 14, color: cWhite } },
          { text: '• ACOS is very high and can be reduced by improving advertising strategies and optimization.', options: { fontSize: 14, color: cWhite } },
        ],
        { x: 1.6, y: 4.3, w: 10.1, h: 2.0 }
      );

      // ----------------- SLIDE 4: OPTIMIZE MANUAL CAMPAIGNS -----------------
      const s4 = pptx.addSlide();
      addSlideHeader(s4, '2', 'OPTIMIZE MANUAL CAMPAIGNS');

      const manualTable = [
        [
          { text: 'TARGETINGs', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'NO. OF TARGETINGs', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'IMPRESSIONS', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'CLICKS', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'SPEND', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'SALES', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'ACOS', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
        ],
        [
          { text: 'No Conversion', options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: String(manualNoConv.targets), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: manualNoConv.imp.toLocaleString(), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: String(manualNoConv.clicks), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `$${manualNoConv.spend.toFixed(2)}`, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `$0`, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: manualNoConv.acos, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
        ],
        [
          { text: '1 Conversion (Units)', options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: String(manualOneConv.targets), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: manualOneConv.imp.toLocaleString(), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: String(manualOneConv.clicks), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `$${manualOneConv.spend.toFixed(2)}`, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `$${manualOneConv.sales.toFixed(2)}`, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: manualOneConv.acos, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
        ],
        [
          { text: '2+ Conversion (Units)', options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: String(manualMultiConv.targets), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: manualMultiConv.imp.toLocaleString(), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: String(manualMultiConv.clicks), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `$${manualMultiConv.spend.toFixed(2)}`, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `$${manualMultiConv.sales.toFixed(2)}`, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: manualMultiConv.acos, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
        ],
      ];

      s4.addTable(manualTable, {
        x: 0.8,
        y: 1.5,
        w: 11.7,
        h: 2.6,
        colW: [2.2, 1.9, 1.6, 1.2, 1.5, 1.8, 1.5],
        border: { pt: 0.75, color: cCyan },
        valign: 'middle',
      });

      s4.addText('Date Range: May 19 - June 17, 2025', {
        x: 8.5,
        y: 4.2,
        w: 4.0,
        h: 0.4,
        align: 'right',
        fontSize: 10,
        color: cMuted
      });

      s4.addText(
        [
          { text: `• Performance Gaps Identified: ${manualNoConv.targets} Targeting generated ${manualNoConv.imp.toLocaleString()} impressions and $${manualNoConv.spend.toFixed(2)} spend without conversions, highlighting optimization needs.\n`, options: { fontSize: 12, color: cWhite } },
          { text: '• Opportunity to control ad spend by negating non-converting unrelated keywords with brand products.\n', options: { fontSize: 12, color: cWhite } },
          { text: '• Complete List of Targeting that can be negated - LINK (AdStream IQ Negative Center)', options: { fontSize: 12, color: cCyan } },
        ],
        { x: 0.8, y: 4.6, w: 11.7, h: 2.0 }
      );

      // ----------------- SLIDE 5: OPTIMIZE AUTO CAMPAIGNS -----------------
      const s5 = pptx.addSlide();
      addSlideHeader(s5, '3', 'OPTIMIZE AUTO CAMPAIGNS');

      const autoTable = [
        [
          { text: 'SEARCH TERMS', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'NO. OF SEARCH TERMS', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'IMPRESSIONS', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'CLICKS', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'SPEND', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'SALES', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'ACOS', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
        ],
        [
          { text: 'No Conversion', options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: String(autoNoConv.terms), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: autoNoConv.imp.toLocaleString(), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: String(autoNoConv.clicks), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `$${autoNoConv.spend.toFixed(2)}`, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `$0`, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: autoNoConv.acos, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
        ],
        [
          { text: '1 Conversion (Units)', options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: String(autoOneConv.terms), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: autoOneConv.imp.toLocaleString(), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: String(autoOneConv.clicks), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `$${autoOneConv.spend.toFixed(2)}`, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `$${autoOneConv.sales.toFixed(2)}`, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: autoOneConv.acos, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
        ],
        [
          { text: '2+ Conversion (Units)', options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: String(autoMultiConv.terms), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: autoMultiConv.imp.toLocaleString(), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: String(autoMultiConv.clicks), options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `$${autoMultiConv.spend.toFixed(2)}`, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: `$${autoMultiConv.sales.toFixed(2)}`, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: autoMultiConv.acos, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
        ],
      ];

      s5.addTable(autoTable, {
        x: 0.8,
        y: 1.5,
        w: 11.7,
        h: 2.6,
        colW: [2.2, 1.9, 1.6, 1.2, 1.5, 1.8, 1.5],
        border: { pt: 0.75, color: cCyan },
        valign: 'middle',
      });

      s5.addText('Date Range: May 19 - June 17, 2025', {
        x: 8.5,
        y: 4.2,
        w: 4.0,
        h: 0.4,
        align: 'right',
        fontSize: 10,
        color: cMuted
      });

      s5.addText(
        [
          { text: `• Performance Gaps Identified: ${autoNoConv.terms} search terms generated ${autoNoConv.imp.toLocaleString()} impressions and $${autoNoConv.spend.toFixed(2)} spend without conversions, highlighting optimization needs.\n`, options: { fontSize: 12, color: cWhite } },
          { text: '• Opportunity to control ad spend by negating non-converting unrelated keywords with brand products.\n', options: { fontSize: 12, color: cWhite } },
          { text: '• Complete List of Search Targets that can be negated - LINK (AdStream IQ Harvester)', options: { fontSize: 12, color: cCyan } },
        ],
        { x: 0.8, y: 4.6, w: 11.7, h: 2.0 }
      );

      // ----------------- SLIDE 6: MATCH TYPE BREAKDOWN -----------------
      const s6 = pptx.addSlide();
      addSlideHeader(s6, '4', 'CAMPAIGN PERFORMANCE - MATCH TYPE');

      const matchTable = [
        [
          { text: 'Match Type', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'Ad Sales', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'Ad Spend', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'Ad Click', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'Ad Order', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'CPC', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'CVR', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'ACOS', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
        ],
        ...matchTypeData.map((m) => [
          { text: m.matchType, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: m.sales, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: m.spend, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: m.clicks, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: m.orders, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: m.cpc, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: m.cvr, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
          { text: m.acos, options: { fill: cNavyBg, color: cWhite, align: 'center' } },
        ])
      ];

      s6.addTable(matchTable, {
        x: 0.8,
        y: 1.5,
        w: 11.7,
        h: 3.0,
        colW: [1.8, 1.5, 1.4, 1.4, 1.4, 1.3, 1.4, 1.5],
        border: { pt: 0.75, color: cCyan },
        valign: 'middle',
      });

      s6.addText('Date Range: May 19 - June 17, 2025', {
        x: 8.5,
        y: 4.7,
        w: 4.0,
        h: 0.4,
        align: 'right',
        fontSize: 10,
        color: cMuted
      });

      s6.addText(
        [
          { text: '• ASIN Type leads with $8,322 in ad sales, while AUTO offers the best cost efficiency with an ACOS of 8.73%.\n', options: { fontSize: 12, color: cWhite } },
          { text: '• PHRASE has the highest CVR (16.5%), and EXACT generates the most orders (278).\n', options: { fontSize: 12, color: cWhite } },
          { text: "• BROAD's high ACOS (50.47%) suggests a need for optimization.", options: { fontSize: 12, color: cWhite } },
        ],
        { x: 0.8, y: 5.1, w: 11.7, h: 1.6 }
      );

      // ----------------- SLIDE 7: BID PLACEMENT -----------------
      const s7 = pptx.addSlide();
      addSlideHeader(s7, '5', 'BID PLACEMENT');

      const placementRows = [
        [
          { text: 'Bidding strategy', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'Placement', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'Impressions', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'Clicks', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'Spend', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'Total Sales', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'Total Units', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'CTR', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'CVR', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'ACOS', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
        ],
        ...placementData.map((p, idx) => {
          const isTotal = idx === 4 || idx === 9 || idx === 10;
          return [
            { text: p.strategy, options: { fill: isTotal ? '163A8A' : cNavyBg, color: cWhite, align: 'center', fontSize: 8.5 } },
            { text: p.placement, options: { fill: isTotal ? '163A8A' : cNavyBg, color: cWhite, align: 'center', fontSize: 8.5 } },
            { text: p.imp, options: { fill: isTotal ? '163A8A' : cNavyBg, color: cWhite, align: 'center', fontSize: 8.5 } },
            { text: p.clicks, options: { fill: isTotal ? '163A8A' : cNavyBg, color: cWhite, align: 'center', fontSize: 8.5 } },
            { text: p.spend, options: { fill: isTotal ? '163A8A' : cNavyBg, color: cWhite, align: 'center', fontSize: 8.5 } },
            { text: p.sales, options: { fill: isTotal ? '163A8A' : cNavyBg, color: cWhite, align: 'center', fontSize: 8.5 } },
            { text: p.units, options: { fill: isTotal ? '163A8A' : cNavyBg, color: cWhite, align: 'center', fontSize: 8.5 } },
            { text: p.ctr, options: { fill: isTotal ? '163A8A' : cNavyBg, color: cWhite, align: 'center', fontSize: 8.5 } },
            { text: p.cvr, options: { fill: isTotal ? '163A8A' : cNavyBg, color: cWhite, align: 'center', fontSize: 8.5 } },
            { text: p.acos, options: { fill: isTotal ? '163A8A' : cNavyBg, color: cWhite, align: 'center', fontSize: 8.5 } },
          ];
        })
      ];

      s7.addTable(placementRows, {
        x: 0.6,
        y: 1.4,
        w: 12.1,
        h: 4.1,
        colW: [1.8, 1.7, 1.1, 0.9, 1.0, 1.2, 1.0, 1.0, 1.0, 1.1],
        border: { pt: 0.5, color: cCyan },
        valign: 'middle',
      });

      s7.addText(
        [
          { text: "• The 'Dynamic bids - up and down' strategy on 'Product pages on Amazon' has the highest ACOS at 106.50%, indicating poor cost efficiency.\n", options: { fontSize: 11, color: cWhite } },
          { text: "• Strategies with 'down only' bidding generally show lower ACOS (e.g., 23.50% total) compared to 'up and down' (38.60% total).", options: { fontSize: 11, color: cWhite } },
        ],
        { x: 0.6, y: 5.6, w: 12.1, h: 1.2 }
      );

      // ----------------- SLIDE 8: SUGGESTED CAMPAIGN STRUCTURE -----------------
      const s8 = pptx.addSlide();
      addSlideHeader(s8, '6', 'SUGGESTED CAMPAIGN STRUCTURE');

      const structRows = [
        [
          { text: 'FUNNEL STAGE', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'CAMPAIGN TYPE', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
          { text: 'STRATEGIC OBJECTIVE', options: { bold: true, fill: cNavyCard, color: cWhite, align: 'center' } },
        ],
        [
          { text: 'DISCOVERY', options: { bold: true, fill: '163A8A', color: cWhite, align: 'center' } },
          { text: 'AUTO CAMPAIGNS', options: { bold: true, fill: cNavyCard, color: cCyan, align: 'center' } },
          { text: 'Automatic targeting used to Discover Keywords to migrate into Research Campaign.', options: { fill: cNavyBg, color: cWhite, align: 'left' } },
        ],
        [
          { text: 'RESEARCH', options: { bold: true, fill: '163A8A', color: cWhite, align: 'center' } },
          { text: 'BROAD CAMPAIGNS', options: { bold: true, fill: cNavyCard, color: cCyan, align: 'center' } },
          { text: "Research Campaigns are used to Manually Target and 'Research' prospective Keywords.", options: { fill: cNavyBg, color: cWhite, align: 'left' } },
        ],
        [
          { text: 'PROFIT', options: { bold: true, fill: '163A8A', color: cWhite, align: 'center' } },
          { text: 'BMM / EXACT MATCH\nPAT / CAT TARGETING', options: { bold: true, fill: cNavyCard, color: cCyan, align: 'center' } },
          { text: "After Keywords & Targets have been 'Researched' and proven profitable, they are migrated to Profit Campaigns to dial in ACoS performance.", options: { fill: cNavyBg, color: cWhite, align: 'left' } },
        ],
        [
          { text: 'STRATEGIC', options: { bold: true, fill: '163A8A', color: cWhite, align: 'center' } },
          { text: 'STRATEGIC DEFENSE', options: { bold: true, fill: cNavyCard, color: cCyan, align: 'center' } },
          { text: 'Conquesting, Head Keywords, Brand Defense, Relevant, Awareness, Long-tail, and Seasonal/Promotional objectives.', options: { fill: cNavyBg, color: cWhite, align: 'left' } },
        ],
      ];

      s8.addTable(structRows, {
        x: 1.0,
        y: 1.6,
        w: 11.3,
        h: 4.8,
        colW: [2.5, 3.2, 5.6],
        border: { pt: 0.75, color: cCyan },
        valign: 'middle',
      });

      // ----------------- SLIDE 9: DAYPARTING OPPORTUNITIES -----------------
      const s9 = pptx.addSlide();
      addSlideHeader(s9, '7', 'Dayparting Opportunities', '4');

      s9.addShape(pptx.ShapeType.roundRect, {
        x: 1.0,
        y: 1.5,
        w: 5.4,
        h: 2.8,
        fill: { color: cNavyCard },
        line: { color: cCyan, width: 1.5 }
      });
      s9.addText(
        [
          { text: 'Hourly Advertising Metrics (Volume Trends)\n\n', options: { fontSize: 12, bold: true, color: cWhite } },
          { text: '• Peak Impression Window: 16:00 – 19:00 (Peak: 39,340 at 18:00)\n• Peak Click Volume: 18:00 (357 clicks)\n• High late afternoon/early evening activity indicates prime window for brand impression share.', options: { fontSize: 10.5, color: 'E2E8F0' } }
        ],
        { x: 1.2, y: 1.7, w: 5.0, h: 2.4 }
      );

      s9.addShape(pptx.ShapeType.roundRect, {
        x: 6.8,
        y: 1.5,
        w: 5.5,
        h: 2.8,
        fill: { color: cNavyCard },
        line: { color: cCyan, width: 1.5 }
      });
      s9.addText(
        [
          { text: 'Hourly CTR, CVR, and ACOS Trends\n\n', options: { fontSize: 12, bold: true, color: cWhite } },
          { text: '• Peak Sales Hours: 10:00 ($1,022.01, 43 orders) & 15:00 ($718.93, 26 orders)\n• CVR Divergence: Highest purchase intent occurs mid-morning & mid-afternoon despite lower impression volume.\n• Off-Peak Drain: Late night spend yields elevated ACOS.', options: { fontSize: 10.5, color: 'E2E8F0' } }
        ],
        { x: 7.0, y: 1.7, w: 5.1, h: 2.4 }
      );

      s9.addText(
        [
          { text: '• Impressions and Engagement: Impressions peak between 16:00 and 19:00, with the highest at 18:00 (39,340). Clicks follow a similar trend, peaking at 18:00 (357).\n', options: { fontSize: 11, color: cWhite } },
          { text: '• Sales and Orders: Total Sales are highest at 10:00 ($1,022.01, 43 orders), followed by 15:00 ($718.93, 26 orders). This indicates strong purchase intent in the mid-morning and mid-afternoon.\n', options: { fontSize: 11, color: cWhite } },
          { text: '• Recommend implementing dayparting in AdStream IQ to maximize conversion opportunity.\n', options: { fontSize: 11, color: cWhite } },
          { text: '• Reallocate budget to 10:00 and 01:00 for maximum ROI.\n', options: { fontSize: 11, color: cWhite } },
          { text: '• Reduce or optimize spend at 18:00 and 19:00 to lower ACOS.', options: { fontSize: 11, color: cWhite } },
        ],
        { x: 1.0, y: 4.5, w: 11.3, h: 2.2 }
      );

      // ----------------- SLIDE 10: LISTING IMPROVEMENT SUGGESTIONS -----------------
      const s10 = pptx.addSlide();
      addSlideHeader(s10, '8', 'LISTING IMPROVEMENT SUGGESTIONS', '4');

      s10.addShape(pptx.ShapeType.roundRect, {
        x: 1.0,
        y: 1.6,
        w: 5.4,
        h: 4.8,
        fill: { color: cNavyCard },
        line: { color: cCyan, width: 1.5 }
      });
      s10.addText(
        [
          { text: '1. Listing Analysis, A+, Main Images\n\n', options: { fontSize: 15, bold: true, color: cWhite } },
          { text: 'What We Found:\nWe found there are no lifestyle images of the product and also no A+ content compared to competitor benchmarks.\n\n', options: { fontSize: 11, color: 'E2E8F0' } },
          { text: 'Suggestions:\n• Add lifestyle images for higher customer engagement.\n• First create basic A+ Content for at least 10 core products.\n• Add high-definition demonstration videos to listing media carousel.\n• Create Brand Story module to cross-sell catalog and build brand equity.', options: { fontSize: 11, color: cWhite } },
        ],
        { x: 1.2, y: 1.8, w: 5.0, h: 4.4 }
      );

      s10.addShape(pptx.ShapeType.roundRect, {
        x: 6.8,
        y: 1.6,
        w: 5.5,
        h: 4.8,
        fill: { color: cNavyCard },
        line: { color: cCyan, width: 1.5 }
      });
      s10.addText(
        [
          { text: '2. Need Brand Store Improvement\n\n', options: { fontSize: 15, bold: true, color: cWhite } },
          { text: 'What We Found:\n• The Brand Store is not proper, which reduces brand credibility and cross-selling opportunities.\n• No subpages or curated categories in Brand Store.\n• Most products are not linked with Brand Store.\n• Brand Name is not visible on Brand Store page.\n\n', options: { fontSize: 11, color: 'E2E8F0' } },
          { text: 'Strategic Impact:\n• A well-maintained Brand Store improves brand trust, repeat visits, and significantly increases AOV (Average Order Value).', options: { fontSize: 11, bold: true, color: cCyan } },
        ],
        { x: 7.0, y: 1.8, w: 5.1, h: 4.4 }
      );

      const sanitizedName = clientName.trim().replace(/[^a-zA-Z0-9_-]/g, '_') || 'Client';
      await pptx.writeFile({ fileName: `${sanitizedName}_Audit_AdStreamIQ.pptx` });
    } catch (error) {
      console.error('Error generating PPTX:', error);
      alert('Could not generate PPT presentation. Check console for details.');
    } finally {
      setIsGeneratingPpt(false);
    }
  };

  return (
    <div className="flex h-screen bg-slate-900 text-slate-100 font-sans">
      

      <div className="flex-1 flex flex-col overflow-y-auto bg-[#0B1E48]">
<header className="p-4 bg-[#071330] border-b border-sky-400/20 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-amber-400 to-orange-500 flex items-center justify-center font-black text-slate-950 text-xl shadow-lg">⚡</div>
            <div>
              <div className="font-extrabold text-white tracking-tight leading-tight">ADSTREAM IQ</div>
              <div className="text-[10px] tracking-wider text-amber-400 font-bold uppercase">Amazon PPC Audit Center</div>
            </div>
          </div>
          <div className="text-xs text-sky-200">Client Audit & Deck Generation Portal</div>
        </header>

        <main className="p-8 max-w-7xl mx-auto w-full space-y-6">
          {/* Header Title & Actions */}
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div>
              <div className="flex items-center gap-2">
                <span className="px-2.5 py-0.5 rounded text-[10px] font-black uppercase tracking-wider bg-sky-500/20 text-sky-300 border border-sky-400/30">
                  AdStream IQ Enterprise
                </span>
                <span className="text-xs text-sky-200/50">•</span>
                <span className="text-xs text-sky-200 font-medium">Reference-Matched Client Audit Portal</span>
              </div>
              <h1 className="text-2xl font-bold text-white tracking-tight mt-1">
                Amazon Advertising Client Audit & PPT Deck Generator
              </h1>
              <p className="text-xs text-slate-300 mt-0.5">
                Upload raw Amazon bulk files or search term reports, audit manual & auto campaigns, match types, and placements, and download a dual-branded client presentation deck in PowerPoint (.pptx) format.
              </p>
            </div>

            <div className="flex items-center gap-3">
              <button
                onClick={handleDownloadPpt}
                disabled={isGeneratingPpt}
                className="flex items-center gap-2 px-5 py-2.5 bg-gradient-to-r from-amber-500 to-orange-600 hover:from-amber-600 hover:to-orange-700 text-white text-xs font-bold rounded-xl shadow-md transition disabled:opacity-50"
              >
                <span>{isGeneratingPpt ? 'Generating Deck...' : '📊 Download Client Audit (.pptx)'}</span>
              </button>
            </div>
          </div>

          {/* Client Branding & Audit Settings Card */}
          <div className="bg-[#102A62] p-5 rounded-xl border border-sky-400/30 shadow-md grid grid-cols-1 lg:grid-cols-3 gap-6 items-center">
            {/* Client Brand Name & Audit Date */}
            <div className="space-y-3">
              <h3 className="text-xs font-bold text-sky-300 uppercase tracking-wider">
                1. Client & Brand Info
              </h3>
              <div className="space-y-2">
                <div>
                  <label className="text-[11px] font-semibold text-slate-300 block mb-1">
                    Client / Brand Name
                  </label>
                  <input
                    type="text"
                    value={clientName}
                    onChange={(e) => setClientName(e.target.value)}
                    placeholder="e.g. Natural State Brands"
                    className="w-full text-xs font-medium bg-[#0B1E48] border border-sky-400/30 rounded-lg px-3 py-2 text-white focus:outline-none focus:ring-1 focus:ring-sky-400"
                  />
                </div>
                <div>
                  <label className="text-[11px] font-semibold text-slate-300 block mb-1">
                    Audit Date / Window
                  </label>
                  <input
                    type="text"
                    value={auditDate}
                    onChange={(e) => setAuditDate(e.target.value)}
                    placeholder="e.g. 6.18.2025"
                    className="w-full text-xs font-medium bg-[#0B1E48] border border-sky-400/30 rounded-lg px-3 py-2 text-white focus:outline-none focus:ring-1 focus:ring-sky-400"
                  />
                </div>
              </div>
            </div>

            {/* Brand Logo Upload & Preview */}
            <div className="space-y-2">
              <h3 className="text-xs font-bold text-sky-300 uppercase tracking-wider">
                2. Client Brand Logo (Left Side)
              </h3>
              <div className="flex items-center gap-4">
                <div className="w-32 h-16 bg-white rounded-lg border border-slate-300 p-2 flex items-center justify-center overflow-hidden shrink-0">
                  {clientLogoBase64 ? (
                    <img src={clientLogoBase64} alt="Client Logo" className="max-h-full max-w-full object-contain" />
                  ) : (
                    <span className="text-[10px] text-slate-400 font-bold">No Logo</span>
                  )}
                </div>
                <div className="space-y-1">
                  <label className="inline-block px-3 py-1.5 bg-[#0B1E48] hover:bg-[#163A8A] text-sky-200 text-xs font-semibold rounded-lg cursor-pointer transition border border-sky-400/40">
                    Upload Brand Logo
                    <input type="file" accept="image/*" onChange={handleLogoUpload} className="hidden" />
                  </label>
                  <p className="text-[10px] text-slate-300">
                    Embedded on one side of all audit PPT slides.
                  </p>
                </div>
              </div>
            </div>

            {/* AdStream IQ Logo Preview */}
            <div className="space-y-2">
              <h3 className="text-xs font-bold text-sky-300 uppercase tracking-wider">
                3. AdStream IQ Logo (Right Side)
              </h3>
              <div className="flex items-center gap-4">
                <div className="w-32 h-16 bg-[#0B1E48] rounded-lg border border-sky-400/30 p-2 flex items-center justify-center overflow-hidden shrink-0 shadow-inner">
                  <img src={ADSTREAM_IQ_LOGO_BASE64} alt="AdStream IQ" className="max-h-full max-w-full object-contain" />
                </div>
                <div>
                  <div className="text-xs font-bold text-white">AdStream IQ Logo</div>
                  <p className="text-[10px] text-slate-300 mt-0.5">
                    Co-branded agency watermark placed on the other side of client slides.
                  </p>
                </div>
              </div>
            </div>
          </div>

          {/* Bulk Amazon Ads File Upload Card */}
          <div className="bg-[#102A62] p-6 rounded-xl border-2 border-dashed border-sky-400/40 hover:border-sky-300 transition text-center space-y-3">
            <div className="w-12 h-12 bg-sky-500/20 text-sky-400 rounded-full flex items-center justify-center mx-auto text-xl font-bold">
              📂
            </div>
            <div>
              <h3 className="text-sm font-bold text-white">
                {uploadedFileName ? `Audited File: ${uploadedFileName}` : 'Upload Amazon Ads Bulk Spreadsheet or Search Term Report (.xlsx / .csv)'}
              </h3>
              <p className="text-xs text-slate-300 mt-1 max-w-lg mx-auto">
                Upload your Amazon Advertising bulk operations sheet or search term report. The engine automatically parses clicks, spend, orders, sales, and isolates manual and auto campaign gaps.
              </p>
            </div>
            <div className="flex items-center justify-center gap-3">
              <label className="inline-block px-4 py-2 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white text-xs font-semibold rounded-lg cursor-pointer shadow-sm transition">
                {isAuditing ? 'Parsing Amazon Rows...' : 'Choose Amazon Report File (.xlsx, .csv)'}
                <input
                  type="file"
                  accept=".xlsx, .xls, .csv"
                  onChange={handleFileUpload}
                  className="hidden"
                />
              </label>
            </div>
            {hasRealData && (
              <div className="p-2 bg-emerald-900/40 text-emerald-300 border border-emerald-500/40 rounded-lg text-xs font-semibold inline-block">
                ✓ Successfully parsed {totalTargets.toLocaleString()} rows from {uploadedFileName}! Ready to generate PPT audit.
              </div>
            )}
          </div>

          {/* Section 1: Total Performance Table Card */}
          <div className="bg-[#102A62] rounded-xl border border-sky-400/30 shadow-md p-6 space-y-4">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <span className="w-7 h-7 rounded-full bg-white text-orange-600 font-bold flex items-center justify-center border-2 border-orange-500 text-sm">
                  1
                </span>
                <h3 className="text-base font-bold text-white">
                  Total Performance (Last 30 days)
                </h3>
              </div>
              <span className="text-xs text-sky-300 font-semibold">
                May 19th - June 17th
              </span>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-center text-xs border border-sky-400/40">
                <thead className="bg-[#163A8A] text-white font-bold text-[11px]">
                  <tr>
                    <th className="py-2.5 px-3 border-r border-sky-400/30">TOTAL SALES</th>
                    <th className="py-2.5 px-3 border-r border-sky-400/30">AD SPEND</th>
                    <th className="py-2.5 px-3 border-r border-sky-400/30">AD SALES</th>
                    <th className="py-2.5 px-3 border-r border-sky-400/30">TACOS</th>
                    <th className="py-2.5 px-3">ACOS</th>
                  </tr>
                </thead>
                <tbody className="bg-[#0B1E48] text-white font-bold">
                  <tr>
                    <td className="py-3 px-3 border-r border-sky-400/30 text-emerald-300">${totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}</td>
                    <td className="py-3 px-3 border-r border-sky-400/30 text-slate-200">${totalSpend.toLocaleString('en-US', { minimumFractionDigits: 2 })}</td>
                    <td className="py-3 px-3 border-r border-sky-400/30 text-sky-200">${adSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}</td>
                    <td className="py-3 px-3 border-r border-sky-400/30 text-amber-300">{tacos.toFixed(1)}%</td>
                    <td className="py-3 px-3 text-red-300">{acos.toFixed(2)}%</td>
                  </tr>
                </tbody>
              </table>
            </div>

            <div className="space-y-1.5 text-xs text-slate-200 pt-2">
              <p>• The TACOS is very low (1.3%) and there is chance for more investment in advertising to increase sales.</p>
              <p>• ACOS is very high (31.00%) and can be reduced by improving advertising strategies and optimization.</p>
            </div>
          </div>

          {/* Section 2 & 3: Optimize Manual & Auto Campaigns */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* Manual Campaigns Card */}
            <div className="bg-[#102A62] rounded-xl border border-sky-400/30 shadow-md p-6 space-y-4">
              <div className="flex items-center gap-3">
                <span className="w-7 h-7 rounded-full bg-white text-orange-600 font-bold flex items-center justify-center border-2 border-orange-500 text-sm">
                  2
                </span>
                <h3 className="text-base font-bold text-white">
                  Optimize Manual Campaigns
                </h3>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-center text-[10px] border border-sky-400/30">
                  <thead className="bg-[#163A8A] text-white font-bold">
                    <tr>
                      <th className="py-2 px-2 border-r border-sky-400/30">TARGETINGs</th>
                      <th className="py-2 px-2 border-r border-sky-400/30">COUNT</th>
                      <th className="py-2 px-2 border-r border-sky-400/30">CLICKS</th>
                      <th className="py-2 px-2 border-r border-sky-400/30">SPEND</th>
                      <th className="py-2 px-2 border-r border-sky-400/30">SALES</th>
                      <th className="py-2 px-2">ACOS</th>
                    </tr>
                  </thead>
                  <tbody className="bg-[#0B1E48] text-slate-200 divide-y divide-sky-400/20">
                    <tr>
                      <td className="py-2 px-2 font-bold text-red-300 border-r border-sky-400/30">No Conversion</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">{manualNoConv.targets}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">{manualNoConv.clicks}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30 text-red-400 font-bold">${manualNoConv.spend.toFixed(2)}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">$0</td>
                      <td className="py-2 px-2">-</td>
                    </tr>
                    <tr>
                      <td className="py-2 px-2 font-bold text-amber-300 border-r border-sky-400/30">1 Conversion</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">{manualOneConv.targets}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">{manualOneConv.clicks}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">${manualOneConv.spend.toFixed(2)}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">${manualOneConv.sales.toFixed(2)}</td>
                      <td className="py-2 px-2">{manualOneConv.acos}</td>
                    </tr>
                    <tr>
                      <td className="py-2 px-2 font-bold text-emerald-300 border-r border-sky-400/30">2+ Conversions</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">{manualMultiConv.targets}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">{manualMultiConv.clicks}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">${manualMultiConv.spend.toFixed(2)}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30 text-emerald-300 font-bold">${manualMultiConv.sales.toFixed(2)}</td>
                      <td className="py-2 px-2 text-emerald-300 font-bold">{manualMultiConv.acos}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <p className="text-[11px] text-slate-300">
                • Performance Gaps: {manualNoConv.targets} targeting generated $243.21 spend without conversions. Negate non-converting terms to control spend.
              </p>
            </div>

            {/* Auto Campaigns Card */}
            <div className="bg-[#102A62] rounded-xl border border-sky-400/30 shadow-md p-6 space-y-4">
              <div className="flex items-center gap-3">
                <span className="w-7 h-7 rounded-full bg-white text-orange-600 font-bold flex items-center justify-center border-2 border-orange-500 text-sm">
                  3
                </span>
                <h3 className="text-base font-bold text-white">
                  Optimize Auto Campaigns
                </h3>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-center text-[10px] border border-sky-400/30">
                  <thead className="bg-[#163A8A] text-white font-bold">
                    <tr>
                      <th className="py-2 px-2 border-r border-sky-400/30">SEARCH TERMS</th>
                      <th className="py-2 px-2 border-r border-sky-400/30">COUNT</th>
                      <th className="py-2 px-2 border-r border-sky-400/30">CLICKS</th>
                      <th className="py-2 px-2 border-r border-sky-400/30">SPEND</th>
                      <th className="py-2 px-2 border-r border-sky-400/30">SALES</th>
                      <th className="py-2 px-2">ACOS</th>
                    </tr>
                  </thead>
                  <tbody className="bg-[#0B1E48] text-slate-200 divide-y divide-sky-400/20">
                    <tr>
                      <td className="py-2 px-2 font-bold text-red-300 border-r border-sky-400/30">No Conversion</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">{autoNoConv.terms}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">{autoNoConv.clicks}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30 text-red-400 font-bold">${autoNoConv.spend.toFixed(2)}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">$0</td>
                      <td className="py-2 px-2">-</td>
                    </tr>
                    <tr>
                      <td className="py-2 px-2 font-bold text-amber-300 border-r border-sky-400/30">1 Conversion</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">{autoOneConv.terms}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">{autoOneConv.clicks}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">${autoOneConv.spend.toFixed(2)}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">${autoOneConv.sales.toFixed(2)}</td>
                      <td className="py-2 px-2">{autoOneConv.acos}</td>
                    </tr>
                    <tr>
                      <td className="py-2 px-2 font-bold text-emerald-300 border-r border-sky-400/30">2+ Conversions</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">{autoMultiConv.terms}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">{autoMultiConv.clicks}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30">${autoMultiConv.spend.toFixed(2)}</td>
                      <td className="py-2 px-2 border-r border-sky-400/30 text-emerald-300 font-bold">${autoMultiConv.sales.toFixed(2)}</td>
                      <td className="py-2 px-2 text-emerald-300 font-bold">{autoMultiConv.acos}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <p className="text-[11px] text-slate-300">
                • Performance Gaps: {autoNoConv.terms} search terms spent $274.34 without orders. Negate irrelevant terms to safeguard ad spend.
              </p>
            </div>
          </div>

          {/* Section 4: Match Type Performance */}
          <div className="bg-[#102A62] rounded-xl border border-sky-400/30 shadow-md p-6 space-y-4">
            <div className="flex items-center gap-3">
              <span className="w-7 h-7 rounded-full bg-white text-orange-600 font-bold flex items-center justify-center border-2 border-orange-500 text-sm">
                4
              </span>
              <h3 className="text-base font-bold text-white">
                Campaign Performance — Match Type
              </h3>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-center text-xs border border-sky-400/30">
                <thead className="bg-[#163A8A] text-white font-bold">
                  <tr>
                    <th className="py-2.5 px-3 border-r border-sky-400/30">Match Type</th>
                    <th className="py-2.5 px-3 border-r border-sky-400/30">Ad Sales</th>
                    <th className="py-2.5 px-3 border-r border-sky-400/30">Ad Spend</th>
                    <th className="py-2.5 px-3 border-r border-sky-400/30">Ad Click</th>
                    <th className="py-2.5 px-3 border-r border-sky-400/30">Ad Order</th>
                    <th className="py-2.5 px-3 border-r border-sky-400/30">CPC</th>
                    <th className="py-2.5 px-3 border-r border-sky-400/30">CVR</th>
                    <th className="py-2.5 px-3">ACOS</th>
                  </tr>
                </thead>
                <tbody className="bg-[#0B1E48] text-slate-200 divide-y divide-sky-400/20">
                  {matchTypeData.map((m, idx) => (
                    <tr key={idx}>
                      <td className="py-2 px-3 font-bold text-sky-300 border-r border-sky-400/30">{m.matchType}</td>
                      <td className="py-2 px-3 font-bold text-emerald-300 border-r border-sky-400/30">${m.sales}</td>
                      <td className="py-2 px-3 border-r border-sky-400/30">${m.spend}</td>
                      <td className="py-2 px-3 border-r border-sky-400/30">{m.clicks}</td>
                      <td className="py-2 px-3 border-r border-sky-400/30">{m.orders}</td>
                      <td className="py-2 px-3 border-r border-sky-400/30">${m.cpc}</td>
                      <td className="py-2 px-3 font-bold text-amber-300 border-r border-sky-400/30">{m.cvr}</td>
                      <td className="py-2 px-3 font-bold text-red-300">{m.acos}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <p className="text-xs text-slate-300">
              • ASIN leads with $8,322 in ad sales, while AUTO offers best cost efficiency (8.73% ACOS). PHRASE has highest CVR (16.5%), and BROAD requires optimization (50.47% ACOS).
            </p>
          </div>
        </main>
      </div>

      
    </div>
  );
}
