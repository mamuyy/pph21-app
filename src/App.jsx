import React, { useState, useRef, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const rp  = (n) => "Rp " + Math.abs(Math.round(n||0)).toLocaleString("id-ID");
const rpSigned = (n) => {
  const value = Math.round(n || 0);
  const prefix = value < 0 ? "-Rp " : "Rp ";
  return prefix + Math.abs(value).toLocaleString("id-ID");
};
const num = (v) => parseFloat(v)||0;
const parseSupportOverride = (v) => {
  const raw = String(v ?? "").trim().toLowerCase();
  if (!raw) return null;
  if (["1", "y", "yes", "ya", "true", "tunjang", "ditunjang", "support"].includes(raw)) return true;
  if (["0", "n", "no", "tidak", "false", "tdk", "tidak ditunjang", "nontunjang"].includes(raw)) return false;
  return null;
};

// ─── TER PMK 168/2023 ─────────────────────────────────────────────────────────
const TER_A = [
  [5400000, 0], [5650000, 0.0025], [5950000, 0.005], [6300000, 0.0075], [6750000, 0.01],
  [7500000, 0.0125], [8550000, 0.015], [9650000, 0.0175], [10050000, 0.02], [10350000, 0.0225],
  [10700000, 0.025], [11050000, 0.03], [11600000, 0.035], [12500000, 0.04], [13750000, 0.05],
  [15100000, 0.06], [16950000, 0.07], [19750000, 0.08], [24150000, 0.09], [26450000, 0.10],
  [28000000, 0.11], [30050000, 0.12], [32400000, 0.13], [35400000, 0.14], [39100000, 0.15],
  [43850000, 0.16], [47800000, 0.17], [51400000, 0.18], [56300000, 0.19], [62200000, 0.20],
  [68600000, 0.21], [77500000, 0.22], [89000000, 0.23], [103000000, 0.24], [125000000, 0.25],
  [157000000, 0.26], [206000000, 0.27], [337000000, 0.28], [454000000, 0.29], [550000000, 0.30],
  [695000000, 0.31], [910000000, 0.32], [1400000000, 0.33], [Infinity, 0.34]
];
const TER_B = [
  [6200000, 0], [6500000, 0.0025], [6850000, 0.005], [7300000, 0.0075], [9200000, 0.01],
  [10750000, 0.015], [11250000, 0.02], [11600000, 0.025], [12600000, 0.03], [13600000, 0.04],
  [14950000, 0.05], [16400000, 0.06], [18450000, 0.07], [21850000, 0.08], [26000000, 0.09],
  [27700000, 0.10], [29350000, 0.11], [31450000, 0.12], [33950000, 0.13], [37100000, 0.14],
  [41100000, 0.15], [45800000, 0.16], [49500000, 0.17], [53800000, 0.18], [58500000, 0.19],
  [64000000, 0.20], [71000000, 0.21], [80000000, 0.22], [93000000, 0.23], [109000000, 0.24],
  [129000000, 0.25], [163000000, 0.26], [211000000, 0.27], [374000000, 0.28], [459000000, 0.29],
  [555000000, 0.30], [704000000, 0.31], [957000000, 0.32], [1405000000, 0.33], [Infinity, 0.34]
];
const TER_C = [
  [6600000, 0], [6950000, 0.0025], [7350000, 0.005], [7800000, 0.0075], [8850000, 0.01],
  [9800000, 0.0125], [10950000, 0.015], [11200000, 0.0175], [12050000, 0.02], [12950000, 0.03],
  [14150000, 0.04], [15550000, 0.05], [17050000, 0.06], [19500000, 0.07], [22700000, 0.08],
  [26600000, 0.09], [28100000, 0.10], [30100000, 0.11], [32600000, 0.12], [35400000, 0.13],
  [38900000, 0.14], [43000000, 0.15], [47400000, 0.16], [51200000, 0.17], [55800000, 0.18],
  [60400000, 0.19], [66700000, 0.20], [74500000, 0.21], [83200000, 0.22], [95600000, 0.23],
  [110000000, 0.24], [134000000, 0.25], [169000000, 0.26], [221000000, 0.27], [390000000, 0.28],
  [463000000, 0.29], [561000000, 0.30], [709000000, 0.31], [965000000, 0.32], [1419000000, 0.33],
  [Infinity, 0.34]
];
const TER_CATEGORY_MAP = {
  TK0: TER_A, TK1: TER_A, K0: TER_A,
  TK2: TER_B, TK3: TER_B, K1: TER_B, K2: TER_B,
  K3: TER_C
};
const getTER = (gross, status) => {
  const table = TER_CATEGORY_MAP[status] || TER_A;
  for (const [maxGross, rate] of table) {
    if (gross <= maxGross) return rate;
  }
  return 0.34;
};
const PTKP_MAP = {
  TK0: 54000000, TK1: 58500000, TK2: 63000000, TK3: 67500000,
  K0: 58500000, K1: 63000000, K2: 67500000, K3: 72000000
};

function calcArt17(pkp) {
  let tax = 0, remaining = Math.max(0, Math.floor(pkp / 1000) * 1000);
  const tiers = [[60000000, 0.05], [190000000, 0.15], [250000000, 0.25], [4500000000, 0.30], [Infinity, 0.35]];
  for (const [limit, rate] of tiers) {
    const chunk = Math.min(remaining, limit);
    tax += chunk * rate;
    remaining -= chunk;
    if (remaining <= 0) break;
  }
  return tax;
}

function hitungSatu(emp, policy, mode = "monthly") {
  const isAnnual = mode === "annual";
  const isFinal = mode === "final";
  const bulan = Math.max(1, num(emp.bulanKerja || 12));
  const annualSheetTunjanganPph = Math.max(0, num(emp.tunjanganPphSheet));
  const policyUsed = {
    gajiPokok: parseSupportOverride(emp.gajiPokokD) ?? policy.gajiPokok,
    tunjangan: parseSupportOverride(emp.tunjanganD) ?? policy.tunjangan,
    lembur: parseSupportOverride(emp.lemburD) ?? policy.lembur,
    bonus: parseSupportOverride(emp.bonusD) ?? policy.bonus
  };
  const regularMonthly = [
    {key:"gajiPokok",val:num(emp.gajiPokok),d:policyUsed.gajiPokok},
    {key:"tunjangan",val:num(emp.tunjangan),d:policyUsed.tunjangan},
    {key:"lembur",   val:num(emp.lembur),   d:policyUsed.lembur},
  ];
  const bonus = { key:"bonus", val:num(emp.bonus), d:policyUsed.bonus };
  const monthlyBruto = regularMonthly.reduce((a,b)=>a+b.val,0) + bonus.val;
  const monthlyBrutoD = regularMonthly.reduce((a,b)=>a+(b.d?b.val:0),0) + (bonus.d ? bonus.val : 0);
  const annualBrutoBase = (regularMonthly.reduce((a,b)=>a+b.val,0) * bulan) + bonus.val + annualSheetTunjanganPph;
  const annualBrutoDBase = (regularMonthly.reduce((a,b)=>a+(b.d?b.val:0),0) * bulan) + (bonus.d ? bonus.val : 0) + annualSheetTunjanganPph;
  const bruto = isAnnual ? annualBrutoBase : monthlyBruto;
  const brutoD = isAnnual ? annualBrutoDBase : monthlyBrutoD;
  const rasioBulanan = monthlyBruto > 0 ? monthlyBrutoD / monthlyBruto : 1;
  const rasioTahunan = annualBrutoBase > 0 ? annualBrutoDBase / annualBrutoBase : rasioBulanan;
  const rasio = isAnnual || isFinal ? rasioTahunan : rasioBulanan;
  const bpjsMonthly = num(emp.bpjs);
  const bpjs = isAnnual ? bpjsMonthly * bulan : bpjsMonthly;
  const pphSebelumnya = num(emp.pphSebelumnya);
  
  let tunj = 0, pphTotal = 0, neto = 0, netoAnnual = 0, pkp = 0, ter = 0, annualPphTotal = 0;

  if (!isAnnual && !isFinal) {
    for(let i=0;i<20;i++){
      const brutoWithTunj = bruto + tunj;
      ter = getTER(brutoWithTunj, emp.status);
      const nt = (brutoWithTunj * ter) * rasio;
      if(Math.abs(nt-tunj)<1){tunj=nt;break;} tunj=nt;
    }
    const brutoWithTunj = bruto + tunj;
    ter = getTER(brutoWithTunj, emp.status);
    pphTotal = brutoWithTunj * ter;
    neto = brutoWithTunj - bpjs;
  } else {
    if (annualSheetTunjanganPph > 0) {
      const totalBruto = annualBrutoBase;
      const biayaJab = Math.min(500000 * bulan, totalBruto * 0.05); // Cap disesuaikan masa kerja
      const netoSetahun = totalBruto - biayaJab - (bpjsMonthly * bulan);
      const ptkp = PTKP_MAP[emp.status] || 54000000;
      netoAnnual = netoSetahun;
      pkp = Math.max(0, netoSetahun - ptkp);
      annualPphTotal = calcArt17(pkp);
      tunj = isFinal ? annualPphTotal - pphSebelumnya : annualSheetTunjanganPph;
    } else {
      for(let i=0;i<20;i++){
        const totalBruto = annualBrutoBase + tunj;
        const biayaJab = Math.min(500000 * bulan, totalBruto * 0.05); // Cap disesuaikan masa kerja
        const netoSetahun = totalBruto - biayaJab - (bpjsMonthly * bulan);
        const ptkp = PTKP_MAP[emp.status] || 54000000;
        netoAnnual = netoSetahun;
        pkp = Math.max(0, netoSetahun - ptkp);
        annualPphTotal = calcArt17(pkp);
        const basisPph = isFinal ? annualPphTotal - pphSebelumnya : annualPphTotal;
        const nt = basisPph * rasio;
        if(Math.abs(nt-tunj)<1){tunj=nt;break;} tunj=nt;
      }
    }
    annualPphTotal = calcArt17(pkp);
    pphTotal = isFinal ? annualPphTotal - pphSebelumnya : annualPphTotal;
    neto = monthlyBruto + tunj - bpjsMonthly;
  }

  const pphP = pphTotal * rasio, pphK = pphTotal * (1 - rasio);
  return {
    ...emp,
    bruto, brutoD, brutoT: bruto - brutoD,
    brutoBulanan: monthlyBruto,
    brutoSetahun: annualBrutoBase,
    bpjs, bpjsBulanan: bpjsMonthly, tunj, annualTunjanganPph: annualSheetTunjanganPph, neto, netoAnnual, pkp, ter, pphTotal, annualPphTotal, pphSebelumnya, pphP, pphK,
    takehome: (isAnnual ? bruto : monthlyBruto) - pphK - (isAnnual ? bpjs : bpjsMonthly),
    beban: (isAnnual ? bruto : monthlyBruto) + pphP,
    rasio: rasio * 100,
    policyUsed,
    isAnnual,
    isFinal
  };
}

// ─── EXCEL EXPORT ─────────────────────────────────────────────────────────────
function exportExcel(results, totals, company, period, policy) {
  const wb = XLSX.utils.book_new();

  // ── Sheet 1: REKAP BULANAN
  const rekap = [
    [`REKAP PPh 21 BULANAN — ${company.toUpperCase()}`],
    [`Periode: ${period}  |  Metode: TER (PMK 168/2023)  |  Dicetak: ${new Date().toLocaleDateString("id-ID")}`],
    [],
    ["No","Nama Karyawan","Status PTKP","Gaji Pokok","Tunjangan","Lembur","Bonus","BPJS","Bruto Total","Tunjangan PPh","Neto Basis","TER (%)","PPh 21 Total","PPh Karyawan (Potong)","PPh Perusahaan (Tunjang)","Take Home Pay","Beban Perusahaan"],
    ...results.map((r,i)=>[
      i+1, r.nama, r.status,
      num(r.gajiPokok), num(r.tunjangan), num(r.lembur), num(r.bonus), num(r.bpjs),
      r.bruto, r.tunj, r.neto, (r.ter*100).toFixed(2)+"%",
      r.pphTotal, r.pphK, r.pphP, r.takehome, r.beban
    ]),
    [],
    ["","","TOTAL","","","","","",
      totals.bruto,"",totals.neto,"",totals.pphTotal,totals.pphK,totals.pphP,totals.takehome,totals.beban
    ],
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(rekap);
  // Column widths
  ws1["!cols"] = [5,22,10,14,14,12,12,12,14,14,14,8,14,20,22,16,18].map(w=>({wch:w}));
  ws1["!merges"] = [{s:{r:0,c:0},e:{r:0,c:16}},{s:{r:1,c:0},e:{r:1,c:16}}];
  XLSX.utils.book_append_sheet(wb, ws1, "Rekap Bulanan");

  // ── Sheet 2: SLIP INDIVIDUAL
  const slipData = [
    ["SLIP PPh 21 INDIVIDUAL"],
    [`Periode: ${period}  |  ${company}`],
    [],
  ];
  results.forEach((r,i) => {
    slipData.push([`${i+1}. ${r.nama} (${r.status})`]);
    slipData.push(["Komponen","Nilai","Keterangan"]);
    slipData.push(["Gaji Pokok",  num(r.gajiPokok), r.policyUsed?.gajiPokok?"Ditunjang PPh":"Tdk Ditunjang"]);
    slipData.push(["Tunjangan",   num(r.tunjangan), r.policyUsed?.tunjangan?"Ditunjang PPh":"Tdk Ditunjang"]);
    slipData.push(["Lembur",      num(r.lembur),    r.policyUsed?.lembur?"Ditunjang PPh":"Tdk Ditunjang"]);
    slipData.push(["Bonus/THR",   num(r.bonus),     r.policyUsed?.bonus?"Ditunjang PPh":"Tdk Ditunjang"]);
    slipData.push(["BPJS (pengurang)", -num(r.bpjs),"pengurang neto"]);
    slipData.push(["Tunjangan PPh Perusahaan", r.tunj, "objek pajak tambahan"]);
    slipData.push(["Penghasilan Neto (Basis TER)", r.neto, ""]);
    slipData.push([`PPh 21 (TER ${(r.ter*100).toFixed(2)}%)`, r.pphTotal, ""]);
    slipData.push(["  → Ditanggung Perusahaan", r.pphP, "tidak dipotong gaji"]);
    slipData.push(["  → Dipotong dari Karyawan", r.pphK, "mengurangi take home"]);
    slipData.push(["TAKE HOME PAY", r.takehome, ""]);
    slipData.push([]);
  });
  const ws2 = XLSX.utils.aoa_to_sheet(slipData);
  ws2["!cols"] = [{wch:30},{wch:18},{wch:24}];
  XLSX.utils.book_append_sheet(wb, ws2, "Slip Individual");

  // ── Sheet 3: RINGKASAN EKSEKUTIF
  const exec = [
    ["RINGKASAN EKSEKUTIF PPh 21"],
    [`${company}  —  ${period}`],
    [],
    ["Keterangan","Jumlah","Satuan"],
    ["Total Karyawan", results.length, "orang"],
    ["Total Bruto Penghasilan", totals.bruto, "Rp"],
    ["PPh 21 Ditanggung Karyawan", totals.pphK, "Rp"],
    ["PPh 21 Ditanggung Perusahaan", totals.pphP, "Rp"],
    ["Total PPh 21 Disetor ke Kas Negara", totals.pphTotal, "Rp"],
    ["Total Take Home Pay", totals.takehome, "Rp"],
    ["Total Beban Riil Perusahaan", totals.beban, "Rp"],
    [],
    ["Kebijakan Tunjangan PPh",""],
    ["Gaji Pokok", policy.gajiPokok?"DITUNJANG":"TIDAK DITUNJANG"],
    ["Tunjangan",  policy.tunjangan?"DITUNJANG":"TIDAK DITUNJANG"],
    ["Lembur",     policy.lembur?"DITUNJANG":"TIDAK DITUNJANG"],
    ["Bonus/THR",  policy.bonus?"DITUNJANG":"TIDAK DITUNJANG"],
    [],
    ["Referensi Regulasi","PMK 168/2023 — Metode TER"],
    ["Dicetak oleh","Mamuyy PPh 21"],
    ["Tanggal cetak", new Date().toLocaleDateString("id-ID")],
  ];
  const ws3 = XLSX.utils.aoa_to_sheet(exec);
  ws3["!cols"] = [{wch:36},{wch:20},{wch:10}];
  ws3["!merges"] = [{s:{r:0,c:0},e:{r:0,c:2}},{s:{r:1,c:0},e:{r:1,c:2}}];
  XLSX.utils.book_append_sheet(wb, ws3, "Ringkasan Eksekutif");

  // ── Download
  XLSX.writeFile(wb, `PPh21_${company.replace(/\s+/g,"_")}_${period.replace(/\s+/g,"_")}.xlsx`);
}

function exportSettlementExcel(results, company, period) {
  const wb = XLSX.utils.book_new();
  const safeCompany = company.replace(/\s+/g, "_");
  const safePeriod = period.replace(/\s+/g, "_");

  const refundYbsRows = results
    .filter((r) => num(r.pphK) < 0)
    .map((r, i) => ({
      No: i + 1,
      Nama: r.nama,
      Status: r.status,
      "Refund ke Ybs": Math.abs(num(r.pphK)),
      "PPh Masa Terakhir": num(r.pphTotal),
      "PPh Setahun": num(r.annualPphTotal),
      "PPh s.d. Lalu": num(r.pphSebelumnya),
    }));

  const refundPerusahaanRows = results
    .filter((r) => num(r.pphP) < 0)
    .map((r, i) => ({
      No: i + 1,
      Nama: r.nama,
      Status: r.status,
      "Refund ke Perusahaan": Math.abs(num(r.pphP)),
      "PPh Masa Terakhir": num(r.pphTotal),
      "PPh Setahun": num(r.annualPphTotal),
      "PPh s.d. Lalu": num(r.pphSebelumnya),
    }));

  const payableYbsRows = results
    .filter((r) => num(r.pphK) > 0)
    .map((r, i) => ({
      No: i + 1,
      Nama: r.nama,
      Status: r.status,
      "Tagih ke Ybs": num(r.pphK),
      "PPh Masa Terakhir": num(r.pphTotal),
      "PPh Setahun": num(r.annualPphTotal),
      "PPh s.d. Lalu": num(r.pphSebelumnya),
    }));

  const payablePerusahaanRows = results
    .filter((r) => num(r.pphP) > 0)
    .map((r, i) => ({
      No: i + 1,
      Nama: r.nama,
      Status: r.status,
      "Beban Perusahaan": num(r.pphP),
      "PPh Masa Terakhir": num(r.pphTotal),
      "PPh Setahun": num(r.annualPphTotal),
      "PPh s.d. Lalu": num(r.pphSebelumnya),
    }));

  const totalRefundYbs = refundYbsRows.reduce((s, r) => s + num(r["Refund ke Ybs"]), 0);
  const totalRefundPerusahaan = refundPerusahaanRows.reduce((s, r) => s + num(r["Refund ke Perusahaan"]), 0);
  const totalTagihYbs = payableYbsRows.reduce((s, r) => s + num(r["Tagih ke Ybs"]), 0);
  const totalBebanPerusahaan = payablePerusahaanRows.reduce((s, r) => s + num(r["Beban Perusahaan"]), 0);

  const summary = [
    ["REKAP SETTLEMENT INTERNAL PPh 21"],
    [`Perusahaan: ${company}`],
    [`Periode: ${period}`],
    [`Dicetak: ${new Date().toLocaleString("id-ID")}`],
    [],
    ["Keterangan", "Jumlah Pegawai", "Nominal"],
    ["Refund ke Ybs", refundYbsRows.length, totalRefundYbs],
    ["Refund ke Perusahaan", refundPerusahaanRows.length, totalRefundPerusahaan],
    ["Tagih ke Ybs", payableYbsRows.length, totalTagihYbs],
    ["Beban Perusahaan", payablePerusahaanRows.length, totalBebanPerusahaan],
    ["Total Refund", "", totalRefundYbs + totalRefundPerusahaan],
    ["Total Payable", "", totalTagihYbs + totalBebanPerusahaan],
  ];
  const wsSummary = XLSX.utils.aoa_to_sheet(summary);
  wsSummary["!cols"] = [{ wch: 30 }, { wch: 18 }, { wch: 20 }];
  wsSummary["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }];
  XLSX.utils.book_append_sheet(wb, wsSummary, "Ringkasan Settlement");

  const wsYbs = XLSX.utils.json_to_sheet(refundYbsRows.length ? refundYbsRows : [{ Info: "Tidak ada data refund ke ybs." }]);
  wsYbs["!cols"] = [
    { wch: 6 }, { wch: 28 }, { wch: 10 }, { wch: 18 }, { wch: 18 }, { wch: 16 }, { wch: 16 }
  ];
  XLSX.utils.book_append_sheet(wb, wsYbs, "Refund ke Ybs");

  const wsPerusahaan = XLSX.utils.json_to_sheet(refundPerusahaanRows.length ? refundPerusahaanRows : [{ Info: "Tidak ada data refund ke perusahaan." }]);
  wsPerusahaan["!cols"] = [
    { wch: 6 }, { wch: 28 }, { wch: 10 }, { wch: 22 }, { wch: 18 }, { wch: 16 }, { wch: 16 }
  ];
  XLSX.utils.book_append_sheet(wb, wsPerusahaan, "Refund ke Perusahaan");

  const payableCombinedRows = [
    ...payableYbsRows.map((r) => ({
      Jenis: "Tagih ke Ybs",
      No: r.No,
      Nama: r.Nama,
      Status: r.Status,
      Nominal: r["Tagih ke Ybs"],
      "PPh Masa Terakhir": r["PPh Masa Terakhir"],
      "PPh Setahun": r["PPh Setahun"],
      "PPh s.d. Lalu": r["PPh s.d. Lalu"],
    })),
    ...payablePerusahaanRows.map((r) => ({
      Jenis: "Beban Perusahaan",
      No: r.No,
      Nama: r.Nama,
      Status: r.Status,
      Nominal: r["Beban Perusahaan"],
      "PPh Masa Terakhir": r["PPh Masa Terakhir"],
      "PPh Setahun": r["PPh Setahun"],
      "PPh s.d. Lalu": r["PPh s.d. Lalu"],
    })),
  ];
  const wsPayable = XLSX.utils.json_to_sheet(
    payableCombinedRows.length ? payableCombinedRows : [{ Info: "Tidak ada data tagih ke ybs / beban perusahaan." }]
  );
  wsPayable["!cols"] = [
    { wch: 20 }, { wch: 6 }, { wch: 28 }, { wch: 10 }, { wch: 18 }, { wch: 18 }, { wch: 16 }, { wch: 16 }
  ];
  XLSX.utils.book_append_sheet(wb, wsPayable, "Tagih & Beban");

  XLSX.writeFile(wb, `Settlement_PPh21_${safeCompany}_${safePeriod}.xlsx`);
}

// ─── SAMPLE CSV ───────────────────────────────────────────────────────────────
const SAMPLE_CSV = `Nama,StatusPajak,GajiPokok,Tunjangan,Lembur,Bonus,BPJS
Ahmad Fauzi,TK0,5000000,500000,800000,0,100000
Siti Rahayu Putri,K1,7500000,750000,0,1000000,150000
Budi Santoso,K2,6000000,600000,500000,500000,120000
Dewi Lestari,TK1,4500000,450000,300000,0,90000
Rudi Hartono,K0,8500000,850000,1200000,2000000,170000
Rina Wulandari,TK0,3800000,380000,600000,0,76000
Hendra Kurniawan,K3,9500000,950000,0,3000000,190000`;

const STATUS_OPTS = ["TK0", "TK1", "TK2", "TK3", "K0", "K1", "K2", "K3"];
const MONTHS = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
const DEFAULT_POLICY = { gajiPokok: true, tunjangan: true, lembur: false, bonus: false };
const INITIAL_EMPLOYEE_COUNT = 4;
const ALIAS_FIELDS = [
  { key: "nama", label: "Nama Karyawan" },
  { key: "status", label: "Status PTKP" },
  { key: "tanggungan", label: "Tanggungan" },
  { key: "bulanKerja", label: "Bulan Kerja" },
  { key: "gajiPokok", label: "Gaji Pokok / Upah Pokok" },
  { key: "tunjangan", label: "Tunjangan / Allowance / Insentif" },
  { key: "lembur", label: "Lembur / Overtime" },
  { key: "bonus", label: "Bonus / THR / Rapel" },
  { key: "bpjs", label: "BPJS / Iuran Pegawai" },
  { key: "pphSebelumnya", label: "PPh Sudah Dipotong Sebelumnya" },
  { key: "pphExcel", label: "PPh Excel Pembanding" },
  { key: "gajiPokokD", label: "Override Ditunjang Gaji (Opsional)" },
  { key: "tunjanganD", label: "Override Ditunjang Tunjangan (Opsional)" },
  { key: "lemburD", label: "Override Ditunjang Lembur (Opsional)" },
  { key: "bonusD", label: "Override Ditunjang Bonus (Opsional)" }
];
const REQUIRED_IMPORT_KEYS = ["nama", "status", "gajiPokok"];
const DEFAULT_ALIAS_TEXT = {
  nama: "nama, nama karyawan, employee, employee name, pegawai, karyawan",
  status: "status, status pajak, status ptkp, ptkp, tax status, status kawin",
  bulanKerja: "bulan kerja, masa kerja, months worked, jumlah bulan",
  gajiPokok: "gaji pokok, basic salary, basic, upah pokok, salary, gapok",
  tunjangan: "tunjangan, allowance, insentif, incentive, premi, uang hadir, tunjangan site, tunjangan jabatan, transport, meal, uang makan",
  lembur: "lembur, overtime, ot",
  bonus: "bonus, thr, rapel, bonus kinerja, insentif tahunan, jasa produksi",
  bpjs: "bpjs, bpjs tk, bpjs kesehatan, iuran bpjs, potongan bpjs, iuran pegawai",
  tanggungan: "tanggungan, dependent, dependents, jumlah tanggungan",
  pphSebelumnya: "pph sebelumnya, pph sd lalu, pph jan nov, pph terpotong sebelumnya, previous pph, cumulative pph",
  pphExcel: "pph excel, pph pembanding, pph21 excel, pph 21 excel, compare pph",
  gajiPokokD: "ditunjang gaji, gaji ditunjang, flag ditunjang gaji, support gaji, tunjang gaji, gaji",
  tunjanganD: "ditunjang tunjangan, tunjangan ditunjang, flag ditunjang tunjangan, support tunjangan, tunjangan pph",
  lemburD: "ditunjang lembur, lembur ditunjang, flag ditunjang lembur, support lembur, tunjangan lainnya lembur",
  bonusD: "ditunjang bonus, bonus ditunjang, flag ditunjang bonus, support bonus, bonus ditunjang pph"
};
const ALIAS_TEMPLATE_STORAGE_KEY = "pph21-alias-templates-v1";
const DEFAULT_ALIAS_TEMPLATES = {
  standard: {
    label: "Standar Umum",
    aliases: DEFAULT_ALIAS_TEXT
  },
  outsourcing: {
    label: "Outsourcing / Site",
    aliases: {
      ...DEFAULT_ALIAS_TEXT,
      nama: `${DEFAULT_ALIAS_TEXT.nama}, nama tenaga kerja, manpower name`,
      tunjangan: `${DEFAULT_ALIAS_TEXT.tunjangan}, insentif site, allowance project, tunjangan project, tunjangan area, premi kehadiran, uang kehadiran, uang transport, uang makan site`,
      bonus: `${DEFAULT_ALIAS_TEXT.bonus}, rapel gaji, koreksi gaji, jasa service, incentive closing`,
      bpjs: `${DEFAULT_ALIAS_TEXT.bpjs}, pot bpjs, potongan jht, potongan jp, potongan kesehatan`
    }
  },
  a1: {
    label: "Kertas Kerja 1721-A1",
    aliases: DEFAULT_ALIAS_TEXT
  }
};

const newEmp = (id) => ({
  id,
  nama: "",
  status: "TK0",
  gajiPokok: "",
  tunjangan: "",
  lembur: "",
  bonus: "",
  bpjs: "",
  pphSebelumnya: "",
  pphExcel: 0,
  bulanKerja: 12,
  gajiPokokD: "",
  tunjanganD: "",
  lemburD: "",
  bonusD: ""
});

const createInitialEmployees = () =>
  Array.from({ length: INITIAL_EMPLOYEE_COUNT }, (_, idx) => newEmp(idx + 1));

const normalizeHeader = (value) =>
  String(value || "")
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();

function normalizeStatus(value) {
  const raw = String(value || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (STATUS_OPTS.includes(raw)) return raw;
  const compact = raw.replace(/^KAWIN/, "K").replace(/^TK/, "TK");
  if (STATUS_OPTS.includes(compact)) return compact;
  return "TK0";
}

function buildStatusPTKP(statusValue, tanggunganValue) {
  const base = String(statusValue || "").toUpperCase().replace(/[^A-Z]/g, "");
  const tanggungan = Math.max(0, Math.min(3, Math.floor(num(tanggunganValue))));
  const candidate = `${base === "K" ? "K" : "TK"}${tanggungan}`;
  return STATUS_OPTS.includes(candidate) ? candidate : normalizeStatus(statusValue);
}

function findHeaderIndex(headers, candidates) {
  const normalizedCandidates = candidates.map(normalizeHeader);
  return headers.findIndex((header) => normalizedCandidates.includes(normalizeHeader(header)));
}

function findHeaderIndexes(headers, candidates) {
  const normalizedCandidates = candidates.map(normalizeHeader);
  const indexes = [];
  headers.forEach((header, idx) => {
    if (normalizedCandidates.includes(normalizeHeader(header))) indexes.push(idx);
  });
  return indexes;
}

function findHeaderIndexPreferred(headers, candidates) {
  for (const candidate of candidates) {
    const idx = findHeaderIndex(headers, [candidate]);
    if (idx !== -1) return idx;
  }
  return -1;
}

function parseA1Payload(rows) {
  const headerRowIndex = rows.findIndex((row) =>
    Array.isArray(row) &&
    row.some((cell) => normalizeHeader(cell) === "nama pegawai") &&
    row.some((cell) => normalizeHeader(cell) === "upah pokok gaji")
  );
  if (headerRowIndex === -1) return null;

  const headers = rows[headerRowIndex] || [];
  const statusIndexes = findHeaderIndexes(headers, ["Status"]);
  const idx = {
    nama: findHeaderIndex(headers, ["Nama Pegawai"]),
    statusDasar: statusIndexes.length > 1 ? statusIndexes[1] : statusIndexes[0] ?? -1,
    tanggungan: findHeaderIndex(headers, ["Tanggungan"]),
    bulanKerja: findHeaderIndex(headers, ["Lama Bekerja"]),
    gajiPokok: findHeaderIndex(headers, ["Upah Pokok / Gaji"]),
    rapel: findHeaderIndex(headers, ["rapel"]),
    tunjanganGaji: findHeaderIndex(headers, ["Total Tunjangan Gaji"]),
    bpjsKesehatan: findHeaderIndex(headers, ["BPJS Kesehatan"]),
    bpjsKetenagakerjaan: findHeaderIndex(headers, ["BPJS Ketenagakerjaan (JKK + JKM"]),
    tunjanganPphSheet: findHeaderIndex(headers, ["Tunjangan PPh"]),
    lembur: findHeaderIndex(headers, ["Lembur"]),
    loyalitas: findHeaderIndex(headers, ["TUNJANGAN LOYALITAS HO"]),
    thr: findHeaderIndex(headers, ["THR/MUEGENG"]),
    bonus: findHeaderIndex(headers, ["Bonus/ ANNUAL INSENTIF INSENTIF BOX SUSULAN)"]),
    kompensasi: findHeaderIndex(headers, ["Kompensasi PKWT"]),
    cuti: findHeaderIndex(headers, ["Cuti"]),
    kekuranganThr: findHeaderIndex(headers, ["KEKURANGAN THR NATAL 2024"]),
    gaji13: findHeaderIndex(headers, ["GAJI KE 13"]),
    honorPic: findHeaderIndex(headers, ["HONOR PIC ADMIN"]),
    lainLain: findHeaderIndex(headers, ["LAIN-LAIN"]),
    honorBko: findHeaderIndex(headers, ["HONOR BKO"]),
    iuranPensiun: findHeaderIndex(headers, ["Iuran Pensiun atau Iuran THT/JHT"]),
    pphSebelumnya: findHeaderIndexPreferred(headers, ["PPh Pasal 21 Jan sd Nov Sudah Dibayar", "PPh Pasal 21 Sebelumnya"]),
    pphExcel: findHeaderIndexPreferred(headers, ["PPh Terutang", "PPh Pasal 21 Terutang"]),
    flagGajiDitunjang: findHeaderIndexPreferred(headers, ["Gaji", "Gaji Ditunjang", "Flag Gaji Ditunjang"]),
    flagTunjanganDitunjang: findHeaderIndexPreferred(headers, ["Tunjangan PPh", "Tunjangan Ditunjang", "Flag Tunjangan Ditunjang"]),
    flagLemburDitunjang: findHeaderIndexPreferred(headers, ["Lembur Ditunjang", "Flag Lembur Ditunjang"]),
    flagBonusDitunjang: findHeaderIndexPreferred(headers, ["Bonus Ditunjang", "Flag Bonus Ditunjang"]),
  };

  const required = ["nama", "statusDasar", "tanggungan", "gajiPokok"];
  if (required.some((key) => idx[key] === -1)) return null;

  const recognizedHeaders = [
    ["nama", idx.nama],
    ["status", idx.statusDasar],
    ["tanggungan", idx.tanggungan],
    ["bulanKerja", idx.bulanKerja],
    ["gajiPokok", idx.gajiPokok],
    ["tunjangan", idx.tunjanganGaji],
    ["lembur", idx.lembur],
    ["bonus", idx.bonus],
    ["bpjs", idx.bpjsKesehatan],
    ["pphSebelumnya", idx.pphSebelumnya],
    ["pphExcel", idx.pphExcel],
    ["gajiPokokD", idx.flagGajiDitunjang],
    ["tunjanganD", idx.flagTunjanganDitunjang],
    ["lemburD", idx.flagLemburDitunjang],
    ["bonusD", idx.flagBonusDitunjang]
  ]
    .filter(([, index]) => index !== -1)
    .map(([target, index]) => ({ target, source: String(headers[index] || "") }));

  const parsedRows = rows
    .slice(headerRowIndex + 1)
    .filter((row) => Array.isArray(row) && String(row[idx.nama] || "").trim())
    .map((row, i) => {
      const bulanKerja = Math.max(1, num(row[idx.bulanKerja] || 12));
      const gajiPokokSetahun = num(row[idx.gajiPokok]);
      const tunjanganSetahun =
        num(row[idx.tunjanganGaji]) +
        num(row[idx.loyalitas]) +
        num(row[idx.honorPic]) +
        num(row[idx.honorBko]) +
        num(row[idx.bpjsKesehatan]) +
        num(row[idx.bpjsKetenagakerjaan]);
      const lemburSetahun = num(row[idx.lembur]);
      const bpjsSetahun = num(row[idx.iuranPensiun]);
      const tunjangan =
        tunjanganSetahun / bulanKerja;
      const bonus =
        num(row[idx.rapel]) +
        num(row[idx.thr]) +
        num(row[idx.bonus]) +
        num(row[idx.kompensasi]) +
        num(row[idx.cuti]) +
        num(row[idx.kekuranganThr]) +
        num(row[idx.gaji13]) +
        num(row[idx.lainLain]);
      const bpjs = bpjsSetahun / bulanKerja;

      return {
        id: Date.now() + i,
        nama: String(row[idx.nama] || `Karyawan ${i + 1}`),
        status: buildStatusPTKP(row[idx.statusDasar], row[idx.tanggungan]),
        gajiPokok: String(gajiPokokSetahun / bulanKerja),
        tunjangan: String(tunjangan),
        lembur: String(lemburSetahun / bulanKerja),
        bonus: String(bonus),
        bpjs: String(bpjs),
        pphSebelumnya: String(num(row[idx.pphSebelumnya])),
        pphExcel: num(row[idx.pphExcel]),
        bulanKerja,
        sumberA1Annual: true,
        tunjanganPphSheet: num(row[idx.tunjanganPphSheet]),
        gajiPokokD: idx.flagGajiDitunjang !== -1 ? String(row[idx.flagGajiDitunjang] ?? "") : "",
        tunjanganD: idx.flagTunjanganDitunjang !== -1 ? String(row[idx.flagTunjanganDitunjang] ?? "") : "",
        lemburD: idx.flagLemburDitunjang !== -1 ? String(row[idx.flagLemburDitunjang] ?? "") : "",
        bonusD: idx.flagBonusDitunjang !== -1 ? String(row[idx.flagBonusDitunjang] ?? "") : "",
      };
    });

  const unmatchedHeaders = headers
    .filter((header) => String(header || "").trim())
    .filter((header) => !recognizedHeaders.some((item) => item.source === String(header || "")));

  return {
    rows: parsedRows,
    headerRowIndex,
    recognizedHeaders,
    unmatchedHeaders,
    warnings: [],
    sourceType: "a1"
  };
}

function buildAliasLookup(aliasText) {
  const lookup = new Map();
  Object.entries(aliasText).forEach(([target, raw]) => {
    String(raw || "")
      .split(",")
      .map((item) => normalizeHeader(item))
      .filter(Boolean)
      .forEach((alias) => lookup.set(alias, target));
  });
  return lookup;
}

function mapHeaderIndexes(headers, aliasLookup) {
  const mapping = {};
  headers.forEach((header, idx) => {
    const direct = aliasLookup.get(normalizeHeader(header));
    if (direct && mapping[direct] === undefined) {
      mapping[direct] = idx;
    }
  });
  return mapping;
}

function parseFlexiblePayload(rows, aliasLookup) {
  const headerRowIndex = rows.findIndex((row) =>
    Array.isArray(row) && row.some((cell) => aliasLookup.has(normalizeHeader(cell)))
  );
  if (headerRowIndex === -1) {
    return {
      rows: [],
      headerRowIndex: -1,
      recognizedHeaders: [],
      unmatchedHeaders: [],
      warnings: []
    };
  }

  const rawHeaders = rows[headerRowIndex] || [];
  const headerIndexes = mapHeaderIndexes(rawHeaders, aliasLookup);
  const recognizedHeaders = Object.entries(headerIndexes).map(([target, idx]) => ({
    target,
    source: String(rawHeaders[idx] || "")
  }));
  const unmatchedHeaders = rawHeaders
    .filter((header, idx) => {
      const normalized = normalizeHeader(header);
      return normalized && !recognizedHeaders.some((item) => item.source === String(rawHeaders[idx] || ""));
    })
    .map((header) => String(header || ""));
  const warnings = [];

  if (headerIndexes.nama === undefined && headerIndexes.gajiPokok === undefined) {
    return {
      rows: [],
      headerRowIndex,
      recognizedHeaders,
      unmatchedHeaders,
      warnings
    };
  }

  const parsedRows = rows
    .slice(headerRowIndex + 1)
    .filter((row) => Array.isArray(row) && row.some((cell) => String(cell ?? "").trim() !== ""))
    .map((row, i) => {
      const rawStatus = row[headerIndexes.status];
      const status = headerIndexes.tanggungan !== undefined
        ? buildStatusPTKP(rawStatus, row[headerIndexes.tanggungan])
        : normalizeStatus(rawStatus);
      const normalizedRawStatus = String(rawStatus || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
      if (String(rawStatus || "").trim() && status === "TK0" && normalizedRawStatus !== "TK0") {
        warnings.push(`Baris ${i + 1}: status "${rawStatus}" dibaca sebagai TK0`);
      }
      return {
        id: Date.now() + i,
        nama: String(row[headerIndexes.nama] || `Karyawan ${i + 1}`),
        status,
        gajiPokok: String(num(row[headerIndexes.gajiPokok])),
        tunjangan: String(num(row[headerIndexes.tunjangan])),
        lembur: String(num(row[headerIndexes.lembur])),
        bonus: String(num(row[headerIndexes.bonus])),
        bpjs: String(num(row[headerIndexes.bpjs])),
        pphSebelumnya: String(num(row[headerIndexes.pphSebelumnya])),
        pphExcel: num(row[headerIndexes.pphExcel]),
        bulanKerja: Math.max(1, num(row[headerIndexes.bulanKerja] || 12)),
        gajiPokokD: String(row[headerIndexes.gajiPokokD] ?? ""),
        tunjanganD: String(row[headerIndexes.tunjanganD] ?? ""),
        lemburD: String(row[headerIndexes.lemburD] ?? ""),
        bonusD: String(row[headerIndexes.bonusD] ?? "")
      };
    });

  return {
    rows: parsedRows,
    headerRowIndex,
    recognizedHeaders,
    unmatchedHeaders,
    warnings
  };
}

function parseFlexibleRows(rows, aliasLookup) {
  return parseFlexiblePayload(rows, aliasLookup).rows;
}

function parseCSV(txt, aliasText) {
  if (!String(txt || "").trim()) return [];
  const workbook = XLSX.read(txt, { type: "string" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: false });
  return parseFlexibleRows(rows, buildAliasLookup(aliasText));
}

function inspectCSV(txt, aliasText) {
  if (!String(txt || "").trim()) {
    return { rows: [], headerRowIndex: -1, recognizedHeaders: [], unmatchedHeaders: [], warnings: [] };
  }
  const workbook = XLSX.read(txt, { type: "string" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: false });
  return parseFlexiblePayload(rows, buildAliasLookup(aliasText));
}

function importExcelRows(rows, mode, aliasText) {
  const a1Payload = parseA1Payload(rows);
  if (a1Payload?.rows?.length) return a1Payload.rows;
  const flexible = parseFlexibleRows(rows, buildAliasLookup(aliasText));
  if (flexible.length) return flexible;
  return rows
    .slice(2)
    .filter((r) => r[4])
    .map((r, i) => {
      const status = String(r[11] || "TK") + String(r[12] || "0");
      const factor = mode === "annual" ? 1 : 12;
      return {
        id: Date.now() + i,
        nama: r[4] || `Karyawan ${i + 1}`,
        status: normalizeStatus(status),
        gajiPokok: String(Math.round(num(r[33]) / factor)),
        tunjangan: "0",
        lembur: "0",
        bonus: "0",
        bpjs: String(Math.round(num(r[38]) / factor)),
        pphSebelumnya: "0",
        pphExcel: num(r[52]),
        bulanKerja: 12
      };
    });
}

function inspectExcelRows(rows, aliasText) {
  const a1Payload = parseA1Payload(rows);
  if (a1Payload) return a1Payload;
  return parseFlexiblePayload(rows, buildAliasLookup(aliasText));
}

function getSourceEmployees(tab, csvText, emps, aliasText) {
  return tab === "csv"
    ? parseCSV(csvText, aliasText)
    : emps.filter((e) => e.nama || num(e.gajiPokok));
}

function sumResults(results) {
  return {
    bruto: results.reduce((s, r) => s + r.bruto, 0),
    neto: results.reduce((s, r) => s + r.neto, 0),
    pphTotal: results.reduce((s, r) => s + r.pphTotal, 0),
    pphK: results.reduce((s, r) => s + r.pphK, 0),
    pphP: results.reduce((s, r) => s + r.pphP, 0),
    takehome: results.reduce((s, r) => s + r.takehome, 0),
    beban: results.reduce((s, r) => s + r.beban, 0)
  };
}

function buildTaxPositionSummary(results) {
  return {
    annualPph: results.reduce((sum, r) => sum + num(r.annualPphTotal), 0),
    priorWithheld: results.reduce((sum, r) => sum + num(r.pphSebelumnya), 0),
    currentObligation: results.reduce((sum, r) => sum + num(r.pphTotal), 0),
    employeeBurden: results.reduce((sum, r) => sum + num(r.pphK), 0),
    companyBurden: results.reduce((sum, r) => sum + num(r.pphP), 0),
  };
}

function buildSettlementSummary(results) {
  return {
    employeePayable: results.reduce((sum, r) => sum + Math.max(0, num(r.pphK)), 0),
    employeeRefund: results.reduce((sum, r) => sum + Math.max(0, -num(r.pphK)), 0),
    companyPayable: results.reduce((sum, r) => sum + Math.max(0, num(r.pphP)), 0),
    companyRefund: results.reduce((sum, r) => sum + Math.max(0, -num(r.pphP)), 0),
  };
}

function buildSettlementBreakdown(results) {
  const employeeRefundRows = results
    .filter((r) => num(r.pphK) < 0)
    .map((r) => ({
      id: r.id,
      nama: r.nama,
      status: r.status,
      amount: Math.abs(num(r.pphK)),
      pphMasaIni: num(r.pphTotal)
    }))
    .sort((a, b) => b.amount - a.amount);

  const companyRefundRows = results
    .filter((r) => num(r.pphP) < 0)
    .map((r) => ({
      id: r.id,
      nama: r.nama,
      status: r.status,
      amount: Math.abs(num(r.pphP)),
      pphMasaIni: num(r.pphTotal)
    }))
    .sort((a, b) => b.amount - a.amount);

  return {
    employeeRefundRows,
    companyRefundRows,
    employeeRefundCount: employeeRefundRows.length,
    companyRefundCount: companyRefundRows.length
  };
}

function buildComparisonSummary(results) {
  const withExcel = results.filter((r) => num(r.pphExcel) !== 0);
  const compared = withExcel.length;
  const mismatches = withExcel.filter((r) => Math.abs(r.pphTotal - r.pphExcel) > 100);
  const appHigher = withExcel.filter((r) => r.pphTotal - r.pphExcel > 100);
  const appLower = withExcel.filter((r) => r.pphExcel - r.pphTotal > 100);
  const totalDiff = withExcel.reduce((sum, r) => sum + (r.pphTotal - r.pphExcel), 0);
  const topDiffs = [...withExcel]
    .sort((a, b) => Math.abs(b.pphTotal - b.pphExcel) - Math.abs(a.pphTotal - a.pphExcel))
    .slice(0, 10);

  return {
    compared,
    mismatches: mismatches.length,
    appHigher: appHigher.length,
    appLower: appLower.length,
    totalDiff,
    topDiffs
  };
}

function filterComparisonSummary(comparison, search) {
  if (!comparison) return null;
  const s = String(search || "").toLowerCase().trim();
  if (!s) return comparison;
  return {
    ...comparison,
    topDiffs: comparison.topDiffs.filter((r) => r.nama.toLowerCase().includes(s))
  };
}

// ─── APP ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [company,  setCompany]  = useState("PT Pelindo Daya Sejahtera");
  const [month,    setMonth]    = useState(new Date().getMonth());
  const [year,     setYear]     = useState(new Date().getFullYear());
  const [tab,      setTab]      = useState("manual");
  const [policy,   setPolicy]   = useState(DEFAULT_POLICY);
  const [emps,     setEmps]     = useState(createInitialEmployees);
  const [csvText,  setCsvText]  = useState("");
  const [results,  setResults]  = useState(null);
  const [mode, setMode] = useState("monthly");
  const [sortK,    setSortK]    = useState("nama");
  const [sortD,    setSortD]    = useState(1);
  const [dragging, setDragging] = useState(false);
  const [loading,  setLoading]  = useState(false);
  const [exporting,setExporting]= useState(false);
  const [exportingPDF, setExportingPDF] = useState(false);
  const [exportingSettlement, setExportingSettlement] = useState(false);
  const [search,   setSearch]   = useState("");
  const [expandId, setExpandId] = useState(null);
  const [showMapping, setShowMapping] = useState(false);
  const [aliasText, setAliasText] = useState(DEFAULT_ALIAS_TEXT);
  const [aliasTemplates, setAliasTemplates] = useState(DEFAULT_ALIAS_TEMPLATES);
  const [selectedTemplate, setSelectedTemplate] = useState("standard");
  const [lastImportInfo, setLastImportInfo] = useState(null);
  const fileRef = useRef();
  const templateFileRef = useRef();

  const period = `${MONTHS[month]} ${year}`;

  const toggleP = (k) => setPolicy(p=>({...p,[k]:!p[k]}));
  const setEmp  = (id,f,v) => setEmps(e=>e.map(r=>r.id===id?{...r,[f]:v}:r));
  const addEmp  = () => setEmps(e=>[...e,newEmp(Date.now())]);
  const delEmp  = (id) => setEmps(e=>e.filter(r=>r.id!==id));

  const handleFile = (file) => {
    setLoading(true);
    const reader = new FileReader();
    const isExcel = file.name.endsWith(".xlsx") || file.name.endsWith(".xls");
    
    reader.onload = (e) => {
      // Gunakan setTimeout agar UI sempat merender status "Loading"
      setTimeout(() => {
        try {
          if (isExcel) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            const inspected = inspectExcelRows(rows, aliasText);
            const mapped = importExcelRows(rows, mode, aliasText);
            setLastImportInfo({
              type: inspected.sourceType === "a1" ? "excel-a1" : "excel",
              recognizedHeaders: inspected.recognizedHeaders,
              unmatchedHeaders: inspected.unmatchedHeaders,
              warnings: inspected.warnings,
              rowCount: mapped.length
            });
            setEmps(mapped);
            setTab("manual");
            alert(`✅ Berhasil import ${mapped.length} karyawan!`);
          } else {
            setCsvText(e.target.result);
            setLastImportInfo(null);
            setTab("csv");
          }
        } catch (err) {
          alert("Gagal membaca file: " + err.message);
        } finally {
          setLoading(false);
        }
      }, 100);
    };
    
    if (isExcel) reader.readAsArrayBuffer(file);
    else reader.readAsText(file);
  };

  const [progress,  setProgress]  = useState(null);

  const sourceEmployees = useMemo(() => getSourceEmployees(tab, csvText, emps, aliasText), [tab, csvText, emps, aliasText]);
  const csvInspection = useMemo(() => inspectCSV(csvText, aliasText), [csvText, aliasText]);
  const currentInspection = tab === "csv" ? csvInspection : lastImportInfo;
  const missingRequiredMappings = useMemo(() => {
    if (!currentInspection?.recognizedHeaders) return [];
    const recognizedKeys = new Set(currentInspection.recognizedHeaders.map((item) => item.target));
    return REQUIRED_IMPORT_KEYS.filter((key) => !recognizedKeys.has(key));
  }, [currentInspection]);
  const importReady = tab === "manual" ? true : missingRequiredMappings.length === 0 && sourceEmployees.length > 0;

  const hitung = () => {
    if (tab === "csv" && missingRequiredMappings.length) {
      return alert(`Kolom wajib belum termapping: ${missingRequiredMappings.join(", ")}`);
    }
    if(!sourceEmployees.length) return alert("Belum ada data karyawan!");
    
    setLoading(true);
    setProgress(0);
    const total = sourceEmployees.length;
    const chunkSize = 1000;
    let res = [];

    const processChunk = (start) => {
      const end = Math.min(start + chunkSize, total);
      const chunk = sourceEmployees.slice(start, end).map(e => ({...e, ...hitungSatu(e, policy, mode)}));
      res = [...res, ...chunk];
      
      const p = Math.round((end / total) * 100);
      setProgress(p);

      if (end < total) {
        setTimeout(() => processChunk(end), 10);
      } else {
        setResults(res);
        setLoading(false);
        setProgress(null);
        setExpandId(null);
      }
    };

    processChunk(0);
  };

  const sorted = useMemo(() => {
    if (!results) return null;
    return [...results].sort((a,b) => {
      const av = a[sortK] || "", bv = b[sortK] || "";
      return typeof av === "number" ? (av - bv) * sortD : String(av).localeCompare(String(bv)) * sortD;
    });
  }, [results, sortK, sortD]);

  const filtered = useMemo(() => {
    if (!sorted) return null;
    const s = search.toLowerCase();
    return sorted.filter(r => !s || r.nama.toLowerCase().includes(s)).slice(0, 100);
  }, [sorted, search]);

  const doSort = (k) => {if(sortK===k)setSortD(d=>-d);else{setSortK(k);setSortD(-1);}};

  const totals = useMemo(() => (results ? sumResults(results) : null), [results]);
  const comparison = useMemo(() => (results ? buildComparisonSummary(results) : null), [results]);
  const visibleComparison = useMemo(() => filterComparisonSummary(comparison, search), [comparison, search]);
  const taxPosition = useMemo(() => (results ? buildTaxPositionSummary(results) : null), [results]);
  const settlement = useMemo(() => (results ? buildSettlementSummary(results) : null), [results]);
  const settlementBreakdown = useMemo(() => {
    if (!results || mode !== "final") return null;
    const source = buildSettlementBreakdown(results);
    const s = String(search || "").toLowerCase().trim();
    if (!s) return source;
    return {
      ...source,
      employeeRefundRows: source.employeeRefundRows.filter((r) => r.nama.toLowerCase().includes(s)),
      companyRefundRows: source.companyRefundRows.filter((r) => r.nama.toLowerCase().includes(s)),
    };
  }, [results, mode, search]);

  const doExport = () => {
    if(!results||!totals) return;
    setExporting(true);
    setTimeout(()=>{
      exportExcel(results, totals, company, period, policy);
      setExporting(false);
    },300);
  };

  const doExportPDF = () => {
    if (!results || !results.length) return;
    setExportingPDF(true);
    
    // Gunakan setTimeout agar UI loading sempat muncul
    setTimeout(() => {
      try {
        const doc = new jsPDF();
        const pageWidth = doc.internal.pageSize.getWidth();
        
        results.forEach((r, i) => {
          if (i > 0) doc.addPage();
          
          // Header
          doc.setFontSize(16);
          doc.setFont("helvetica", "bold");
          doc.text(company.toUpperCase(), pageWidth / 2, 20, { align: "center" });
          
          doc.setFontSize(12);
          doc.text("SLIP PENGHASILAN & PPh 21 (TER)", pageWidth / 2, 28, { align: "center" });
          
          doc.setFontSize(10);
          doc.setFont("helvetica", "normal");
          doc.text(`Periode: ${period}`, pageWidth / 2, 34, { align: "center" });
          
          doc.line(20, 38, pageWidth - 20, 38);
          
          // Data Karyawan
          doc.setFont("helvetica", "bold");
          doc.text("DATA KARYAWAN", 20, 48);
          doc.setFont("helvetica", "normal");
          doc.text(`Nama: ${r.nama}`, 20, 54);
          doc.text(`Status PTKP: ${r.status}`, 20, 59);
          doc.text(`Masa Kerja: ${r.bulanKerja} bulan`, 20, 64);
          
          // Komponen Table
          const body = [
            ["Gaji Pokok", rp(r.gajiPokok), policy.gajiPokok ? "Ditunjang" : "-"],
            ["Tunjangan", rp(r.tunjangan), policy.tunjangan ? "Ditunjang" : "-"],
            ["Lembur", rp(r.lembur), policy.lembur ? "Ditunjang" : "-"],
            ["Bonus/THR", rp(r.bonus), policy.bonus ? "Ditunjang" : "-"],
            ["BPJS (Pengurang)", rp(-r.bpjs), ""],
            ["Tunjangan PPh (dari Perusahaan)", rp(r.tunj), ""],
            ["", "", ""],
            ["PENGHASILAN NETO (BASIS TER)", rp(r.neto), ""],
            [`PPh 21 (TER ${(r.ter * 100).toFixed(2)}%)`, rp(r.pphTotal), ""],
            ["", "", ""],
            ["PPh 21 DITANGGUNG PERUSAHAAN", rp(r.pphP), ""],
            ["PPh 21 DIPOTONG DARI GAJI", rp(r.pphK), ""],
          ];
          
          autoTable(doc, {
            startY: 72,
            head: [["Komponen", "Nilai", "Keterangan"]],
            body: body,
            theme: "striped",
            headStyles: { fillStyle: "#10b981" },
            margin: { left: 20, right: 20 },
          });
          
          const finalY = doc.lastAutoTable.finalY || 150;
          
          // Take Home Pay Box
          doc.setFillColor(245, 245, 245);
          doc.rect(20, finalY + 10, pageWidth - 40, 20, "F");
          doc.setFont("helvetica", "bold");
          doc.setFontSize(12);
          doc.text("TAKE HOME PAY", 30, finalY + 23);
          doc.text(rp(r.takehome), pageWidth - 30, finalY + 23, { align: "right" });
          
          // Footer
          doc.setFontSize(8);
          doc.setFont("helvetica", "italic");
          doc.text(`Dicetak otomatis oleh Mamuyy PPh 21 App pada ${new Date().toLocaleString("id-ID")}`, 20, doc.internal.pageSize.getHeight() - 10);
        });
        
        doc.save(`Slip_PPh21_${company.replace(/\s+/g,"_")}_${period.replace(/\s+/g,"_")}.pdf`);
      } catch (err) {
        alert("Gagal membuat PDF: " + err.message);
      } finally {
        setExportingPDF(false);
      }
    }, 100);
  };

  const doExportSettlement = () => {
    if (!results || !results.length || mode !== "final") return;
    setExportingSettlement(true);
    setTimeout(() => {
      try {
        exportSettlementExcel(results, company, period);
      } finally {
        setExportingSettlement(false);
      }
    }, 120);
  };

  const downloadSample = () => {
    const blob=new Blob([SAMPLE_CSV],{type:"text/csv"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");a.href=url;a.download="template_karyawan.csv";a.click();
  };

  const setAlias = (key, value) => setAliasText((prev) => ({ ...prev, [key]: value }));
  const applyAliasTemplate = (templateKey) => {
    const template = aliasTemplates[templateKey];
    if (!template) return;
    setSelectedTemplate(templateKey);
    setAliasText(template.aliases);
  };
  const saveCurrentTemplate = () => {
    const rawName = window.prompt("Nama template site/klien:");
    const name = String(rawName || "").trim();
    if (!name) return;
    const key = `custom_${name.toLowerCase().replace(/[^a-z0-9]+/g, "_")}`;
    setAliasTemplates((prev) => ({
      ...prev,
      [key]: {
        label: name,
        aliases: aliasText,
        custom: true
      }
    }));
    setSelectedTemplate(key);
  };
  const deleteCurrentTemplate = () => {
    const template = aliasTemplates[selectedTemplate];
    if (!template?.custom) return;
    setAliasTemplates((prev) => {
      const next = { ...prev };
      delete next[selectedTemplate];
      return next;
    });
    setSelectedTemplate("standard");
    setAliasText(DEFAULT_ALIAS_TEMPLATES.standard.aliases);
  };
  const exportAliasTemplates = () => {
    const payload = {
      exportedAt: new Date().toISOString(),
      templates: aliasTemplates
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "pph21-alias-templates.json";
    a.click();
    URL.revokeObjectURL(url);
  };
  const importAliasTemplates = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const parsed = JSON.parse(String(e.target?.result || "{}"));
        const incoming = parsed.templates && typeof parsed.templates === "object" ? parsed.templates : parsed;
        if (!incoming || typeof incoming !== "object") throw new Error("Format template tidak valid");
        setAliasTemplates((prev) => ({ ...prev, ...incoming }));
        alert("Template alias berhasil diimpor.");
      } catch (err) {
        alert("Gagal mengimpor template: " + err.message);
      }
    };
    reader.readAsText(file);
  };

  useEffect(() => {
    try {
      const stored = window.localStorage.getItem(ALIAS_TEMPLATE_STORAGE_KEY);
      if (!stored) return;
      const parsed = JSON.parse(stored);
      if (parsed && typeof parsed === "object") {
        setAliasTemplates((prev) => ({ ...prev, ...parsed }));
      }
    } catch {}
  }, []);

  useEffect(() => {
    try {
      const customOnly = Object.fromEntries(
        Object.entries(aliasTemplates).filter(([, value]) => value?.custom)
      );
      window.localStorage.setItem(ALIAS_TEMPLATE_STORAGE_KEY, JSON.stringify(customOnly));
    } catch {}
  }, [aliasTemplates]);

  const C = {
    root: {fontFamily: "'Inter', system-ui, -apple-system, sans-serif", background: "#0a0f18", color: "#e2e8f0", minHeight: "100vh", padding: "40px 20px", transition: "all 0.3s ease"},
    hdr: {display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 32},
    logoTxt: {fontSize: 22, fontWeight: 800, color: "#fff", letterSpacing: "-0.5px"},
    configBar: {display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 24, background: "rgba(255,255,255,0.03)", padding: 24, borderRadius: 16, marginBottom: 24, border: "1px solid rgba(255,255,255,0.05)"},
    configGroup: {display: "flex", flexDirection: "column", gap: 10},
    configLabel: {fontSize: 10, fontWeight: 700, color: "#94a3b8", letterSpacing: "1.2px", textTransform: "uppercase"},
    configInput: {background: "rgba(0,0,0,0.3)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 8, padding: "10px 14px", color: "#fff", fontSize: 13, outline: "none", transition: "border-color 0.2s"},
    configSel: {background: "rgba(0,0,0,0.3)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 8, padding: "10px 14px", color: "#fff", fontSize: 13, cursor: "pointer", outline: "none"},
    configArea: {background: "rgba(0,0,0,0.3)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 8, padding: "10px 14px", color: "#fff", minHeight: 80, resize: "vertical", fontSize: 12, lineHeight: 1.5},
    chip: {padding: "8px 16px", borderRadius: 24, border: "1px solid", fontSize: 11, cursor: "pointer", fontWeight: 600, transition: "all 0.2s"},
    smallBtn: {background: "rgba(255,255,255,0.05)", border: "1px solid rgba(255,255,255,0.1)", color: "#f1f5f9", padding: "10px 16px", borderRadius: 8, cursor: "pointer", fontSize: 12, fontWeight: 600, transition: "background 0.2s"},
    body: {background: "rgba(255,255,255,0.02)", padding: 32, borderRadius: 20, border: "1px solid rgba(255,255,255,0.03)"},
    tabRow: {display: "flex", gap: 32, marginBottom: 24},
    tab: {background: "transparent", border: "none", padding: "12px 0", cursor: "pointer", fontWeight: 700, fontSize: 13, color: "#64748b", transition: "color 0.2s"},
    calcBtn: {width: "100%", padding: 18, borderRadius: 12, background: "linear-gradient(135deg, #10b981 0%, #059669 100%)", border: "none", fontWeight: 800, color: "#fff", marginTop: 24, cursor: "pointer", fontSize: 15, boxShadow: "0 10px 15px -3px rgba(16, 185, 129, 0.2)"},
    thead: {display: "flex", background: "rgba(255,255,255,0.05)", padding: "12px 20px", borderRadius: 8, marginBottom: 8},
    th: {fontSize: 11, color: "#94a3b8", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.5px"},
    erow: {display: "flex", alignItems: "center", gap: 12, padding: "8px 20px", borderRadius: 8, transition: "background 0.2s"},
    cell: {background: "rgba(0,0,0,0.3)", border: "1px solid rgba(255,255,255,0.08)", borderRadius: 6, padding: "8px 12px", color: "#fff", fontSize: 13, outline: "none"},
    td0: {fontSize: 13, fontWeight: 600}, 
    xBtn: {background: "rgba(239, 68, 68, 0.1)", border: "none", color: "#f87171", cursor: "pointer", width: 24, height: 24, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16},
    addBtn: {background: "rgba(16, 185, 129, 0.1)", border: "1px solid rgba(16, 185, 129, 0.3)", color: "#34d399", padding: "10px 20px", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 13},
    ghostBtn: {background: "transparent", border: "1px solid rgba(148, 163, 184, 0.3)", color: "#94a3b8", padding: "10px 20px", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 13},
    drop: {border: "2px dashed rgba(16, 185, 129, 0.3)", padding: 48, textAlign: "center", borderRadius: 16, cursor: "pointer", transition: "all 0.3s ease", background: "rgba(16, 185, 129, 0.02)"},
    cards: {display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 20, marginBottom: 32},
    card: {padding: 24, borderRadius: 16, border: "1px solid rgba(255,255,255,0.05)", transition: "transform 0.2s ease"},
    exportBar: {display: "flex", gap: 12, alignItems: "center", padding: 16, background: "rgba(255,255,255,0.03)", borderRadius: 12, border: "1px solid rgba(255,255,255,0.05)"},
    exportBtn: {padding: "10px 20px", borderRadius: 8, border: "none", fontSize: 12, fontWeight: 700, cursor: "pointer", transition: "all 0.2s"},
    infoBox: {background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.06)", borderRadius: 16, padding: 20},
    Th: {padding: "14px 12px", textAlign: "left", color: "#94a3b8", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.5px"},
    Tr: {borderBottom: "1px solid rgba(255,255,255,0.04)"}, 
    Td: {padding: "14px 12px", fontSize: 13},
    sectionTitle: {fontSize: 12, fontWeight: 800, color: "#fff", letterSpacing: "1.5px", textTransform: "uppercase", marginBottom: 16, display: "flex", alignItems: "center", gap: 8}
  };
  const taxFlowFmt = mode === "final" ? rpSigned : rp;
  const burdenFmt = (value) => rp(Math.abs(value));

  return (
    <div style={C.root}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');
        .spin { animation: spin 2s linear infinite; display: inline-block; }
        @keyframes spin { 100% { transform: rotate(360deg); } }
        
        /* Custom Scrollbar */
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: rgba(255,255,255,0.02); }
        ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.1); borderRadius: 10px; }
        ::-webkit-scrollbar-thumb:hover { background: rgba(255,255,255,0.2); }

        /* Responsive Table */
        @media (max-width: 768px) {
          body { padding: 10px !important; }
          .config-bar { gridTemplateColumns: 1fr !important; }
          .card-summary { gridTemplateColumns: 1fr 1fr !important; }
          .hide-mobile { display: none !important; }
        }
      `}</style>
      <div style={C.hdr}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <div>
            <div style={C.logoTxt}>Mamuyy <em style={{fontStyle:"normal",color:"#34d399",fontSize:13}}>PPh 21</em></div>
            <div style={{fontSize:11,color:"#6b7fa3"}}>HR Edition · TER (PMK 168/2023) & Pasal 17</div>
          </div>
        </div>
      </div>

        <div style={C.configBar} className="config-bar">
          <div style={C.configGroup}>
          <div style={C.configLabel}>MODE PERHITUNGAN</div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>setMode("monthly")} style={{...C.chip,background:mode==="monthly"?"#34d399":"transparent",borderColor:mode==="monthly"?"#34d399":"#1a3020",color:mode==="monthly"?"#060c16":"#6b7fa3"}}>BULANAN (TER)</button>
            <button onClick={()=>setMode("final")} style={{...C.chip,background:mode==="final"?"#f59e0b":"transparent",borderColor:mode==="final"?"#f59e0b":"#1a3020",color:mode==="final"?"#060c16":"#6b7fa3"}}>MASA TERAKHIR</button>
            <button onClick={()=>setMode("annual")} style={{...C.chip,background:mode==="annual"?"#93c5fd":"transparent",borderColor:mode==="annual"?"#93c5fd":"#1a3020",color:mode==="annual"?"#060c16":"#6b7fa3"}}>TAHUNAN (PASAL 17)</button>
          </div>
          <div style={{fontSize:10,color:"#6b7fa3",lineHeight:1.4}}>
            {mode==="monthly" && "Input diperlakukan sebagai angka bulanan."}
            {mode==="final" && "Mode masa terakhir menghitung rekonsiliasi PPh akhir tahun: PPh setahun dikurangi PPh Jan-Nov / sebelumnya."}
            {mode==="annual" && "Mode tahunan membaca input apa adanya sebagai basis setahun/masa kerja."}
          </div>
        </div>
        <div style={C.configGroup}>
          <span style={C.configLabel}>PERUSAHAAN</span>
          <input style={C.configInput} value={company} onChange={e=>setCompany(e.target.value)} placeholder="Nama Perusahaan"/>
        </div>
        <div style={C.configGroup}>
          <div style={C.configLabel}>PERIODE</div>
          <div style={{display:"flex",gap:6}}>
            <select style={C.configSel} value={month} onChange={e=>setMonth(+e.target.value)}>
              {MONTHS.map((m,i)=><option key={m} value={i}>{m}</option>)}
            </select>
            <select style={C.configSel} value={year} onChange={e=>setYear(+e.target.value)}>
              {[2024,2025,2026,2027].map(y=><option key={y} value={y}>{y}</option>)}
            </select>
          </div>
        </div>
        <div style={C.configGroup}>
          <span style={C.configLabel}>PPh DITUNJANG PERUSAHAAN</span>
          <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
            {Object.entries({gajiPokok:"Gaji",tunjangan:"Tunjangan",lembur:"Lembur",bonus:"Bonus"}).map(([k,lbl])=>(
              <button key={k} onClick={()=>toggleP(k)} style={{
                ...C.chip,
                background:policy[k]?"rgba(52,211,153,0.15)":"rgba(255,255,255,0.04)",
                borderColor:policy[k]?"rgba(52,211,153,0.4)":"rgba(255,255,255,0.08)",
                color:policy[k]?"#34d399":"#3d4f6e",
              }}>{policy[k]?"✅":"☐"} {lbl}</button>
            ))}
          </div>
        </div>
      </div>

      {/* ── SUMMARY DASHBOARD ── */}
      {results && totals && (
        <div style={C.cards}>
          {[
            ["👥 TOTAL KARYAWAN", results.length + " orang", "#fff", "rgba(255,255,255,0.05)", "rgba(255,255,255,0.1)"],
            ["📦 TOTAL BRUTO", rp(totals.bruto), "#93c5fd", "rgba(147,197,253,0.08)", "rgba(147,197,253,0.2)"],
            ["🔴 TOTAL PPh 21", taxFlowFmt(totals.pphTotal), "#fca5a5", "rgba(252,165,165,0.08)", "rgba(252,165,165,0.2)"],
            ["🟢 DITUNJANG PERUSAHAAN", taxFlowFmt(totals.pphP), "#34d399", "rgba(16,185,129,0.08)", "rgba(16,185,129,0.2)"],
            ...(comparison && comparison.compared > 0 ? [
              ["⚖️ SELISIH VS EXCEL", rpSigned(comparison.totalDiff), Math.abs(comparison.totalDiff) > 1000 ? "#f87171" : "#fbbf24", "rgba(251,191,36,0.08)", "rgba(251,191,36,0.2)"]
            ] : [])
          ].map(([l, v, c, bg, br]) => (
            <div key={l} style={{...C.card, background: bg, borderColor: br}} className="card-summary">
              <div style={{fontSize: 10, fontWeight: 700, color: "#94a3b8", marginBottom: 8, letterSpacing: "0.5px"}}>{l}</div>
              <div style={{fontSize: 20, fontWeight: 900, color: c, letterSpacing: "-0.5px"}}>{v}</div>
            </div>
          ))}
        </div>
      )}

      <div style={{...C.body, marginBottom: 24}}>
        <div 
          onClick={() => setShowMapping(!showMapping)} 
          style={{...C.sectionTitle, cursor: "pointer", justifyContent: "space-between", marginBottom: showMapping ? 16 : 0}}
        >
          <span>⚙️ PENGATURAN LANJUTAN / MAPPING ALIAS PAYROLL</span>
          <span style={{fontSize: 14}}>{showMapping ? "−" : "+"}</span>
        </div>
        
        {showMapping && (
          <>
            <div style={{fontSize: 12, color: "#94a3b8", marginBottom: 20, lineHeight: 1.6}}>
              Konfigurasikan nama kolom payroll Anda agar terbaca otomatis oleh sistem. Pisahkan alias dengan koma.
            </div>
            <div style={{display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap", marginBottom: 24}}>
              <select style={{...C.configSel, minWidth: 240}} value={selectedTemplate} onChange={(e) => applyAliasTemplate(e.target.value)}>
                {Object.entries(aliasTemplates).map(([key, value]) => (
                  <option key={key} value={key}>{value.label}</option>
                ))}
              </select>
              <button onClick={saveCurrentTemplate} style={C.smallBtn}>Simpan Template</button>
              <button onClick={exportAliasTemplates} style={C.smallBtn}>Export JSON</button>
              <button onClick={() => templateFileRef.current?.click()} style={C.smallBtn}>Import JSON</button>
              <input ref={templateFileRef} type="file" accept=".json,application/json" style={{display: "none"}} onChange={(e) => importAliasTemplates(e.target.files?.[0])} />
              {aliasTemplates[selectedTemplate]?.custom && (
                <button onClick={deleteCurrentTemplate} style={{...C.smallBtn, color: "#fca5a5", border: "1px solid rgba(252,165,165,0.2)"}}>Hapus</button>
              )}
            </div>
            <div style={{display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 16}}>
              {ALIAS_FIELDS.map((field) => (
                <div key={field.key} style={C.configGroup}>
                  <div style={C.configLabel}>{field.label}</div>
                  <textarea
                    style={C.configArea}
                    value={aliasText[field.key]}
                    onChange={(e) => setAlias(field.key, e.target.value)}
                    placeholder={`Contoh: ${field.label.toLowerCase()}, gapok, basic...`}
                  />
                </div>
              ))}
            </div>
          </>
        )}
        {lastImportInfo && (
          <div style={{...C.infoBox, marginTop: 14}}>
            <div style={{fontSize:10,fontWeight:700,color:"#6b7fa3",letterSpacing:"1.2px",marginBottom:8}}>
              PREVIEW MAPPING IMPORT TERAKHIR ({lastImportInfo.type.toUpperCase()})
            </div>
            <div style={{fontSize:11,color:"#34d399",marginBottom:8}}>
              {lastImportInfo.rowCount} baris berhasil dibaca
            </div>
            {lastImportInfo.type === "excel-a1" && (
              <div style={{fontSize:11,color:"#93c5fd",marginBottom:8}}>
                File dikenali sebagai format Kertas Kerja 1721-A1 dan diparse dengan mapping khusus.
              </div>
            )}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit, minmax(220px, 1fr))",gap:8}}>
              {lastImportInfo.recognizedHeaders.map((item) => (
                <div key={`${item.target}-${item.source}`} style={{fontSize:11,color:"#e8eaf0",background:"rgba(0,0,0,0.2)",padding:"8px 10px",borderRadius:8}}>
                  <strong style={{color:"#93c5fd"}}>{item.source}</strong> → {item.target}
                </div>
              ))}
            </div>
            {!!lastImportInfo.unmatchedHeaders.length && (
              <div style={{marginTop:10,fontSize:11,color:"#fde68a",lineHeight:1.5}}>
                Belum termapping: {lastImportInfo.unmatchedHeaders.join(", ")}
              </div>
            )}
            {!!lastImportInfo.warnings?.length && (
              <div style={{marginTop:10,fontSize:11,color:"#fca5a5",lineHeight:1.5}}>
                Warning import: {lastImportInfo.warnings.slice(0, 5).join(" | ")}
              </div>
            )}
          </div>
        )}
      </div>

      {/* ── BODY ── */}
      <div style={C.body}>

        {/* Tabs */}
        <div style={C.tabRow}>
          {[["manual","✏️ Input Manual"],["csv","📂 Upload CSV"]].map(([t,lbl])=>(
            <button key={t} onClick={()=>setTab(t)} style={{
              ...C.tab,
              borderBottom:tab===t?"2px solid #34d399":"2px solid transparent",
              color:tab===t?"#34d399":"#3d4f6e",
            }}>{lbl}</button>
          ))}
          <div style={{flex:1,borderBottom:"2px solid rgba(255,255,255,0.06)"}}/>
        </div>

        {/* ── MANUAL ── */}
        {tab==="manual" && (
          <div style={{overflowX:"auto"}}>
            <div style={{minWidth:700}}>
              <div style={C.thead}>
                {["No","Nama Karyawan","Status","Bulan Kerja","Gaji Pokok","Tunjangan","Lembur","Bonus","BPJS","PPh s.d. Lalu",""].map((h,i)=>(
                  <div key={i} style={{...C.th, flex:i===1?2:i===0||i===10?0.4:1, minWidth:i===1?130:i===0?30:i===10?28:80, textAlign:i>3&&i<10?"right":"left"}}>{h}</div>
                ))}
              </div>
              <div style={{maxHeight:400,overflowY:"auto"}}>
                {emps.slice(0, 100).map((emp, i) => (
                  <div key={emp.id} style={{...C.erow,background:i%2===0?"rgba(255,255,255,0.015)":"transparent"}}>
                    <div style={{...C.td0,flex:0.4,minWidth:30,color:"#2a4050"}}>{i+1}</div>
                    <input style={{...C.cell,flex:2,minWidth:130}} placeholder={`Karyawan ${i+1}`} value={emp.nama} onChange={e=>setEmp(emp.id,"nama",e.target.value)}/>
                    <select style={{...C.cell,flex:1,minWidth:72,padding:"6px 4px"}} value={emp.status} onChange={e=>setEmp(emp.id,"status",e.target.value)}>
                      {STATUS_OPTS.map(s=><option key={s}>{s}</option>)}
                    </select>
                    <input type="number" style={{...C.cell,flex:1,minWidth:60,textAlign:"center"}} value={emp.bulanKerja} onChange={e=>setEmp(emp.id,"bulanKerja",e.target.value)}/>
                    {["gajiPokok","tunjangan","lembur","bonus","bpjs","pphSebelumnya"].map(f=>(
                      <input key={f} type="number" style={{...C.cell,flex:1,minWidth:80,textAlign:"right"}} placeholder="0" value={emp[f]} onChange={e=>setEmp(emp.id,f,e.target.value)}/>
                    ))}
                    <button onClick={()=>delEmp(emp.id)} style={C.xBtn}>×</button>
                  </div>
                ))}
              </div>
              {emps.length > 100 && (
                <div style={{textAlign:"center",padding:10,fontSize:11,color:"#6b7fa3",background:"rgba(0,0,0,0.2)"}}>
                  ... Menampilkan 100 dari {emps.length} karyawan (Gunakan "Hitung Semua" untuk memproses seluruhnya) ...
                </div>
              )}
              <div style={{display:"flex",gap:8,marginTop:10}}>
                <button onClick={addEmp} style={C.addBtn}>+ Tambah Baris</button>
                <button onClick={()=>setEmps(createInitialEmployees())} style={C.ghostBtn}>Reset</button>
              </div>
            </div>
          </div>
        )}

        {/* ── CSV ── */}
        {tab==="csv" && (
          <div>
            <div style={{display:"flex",gap:8,marginBottom:12,alignItems:"center",flexWrap:"wrap"}}>
              <button onClick={downloadSample} style={C.sampleBtn}>⬇ Download Template CSV</button>
              <span style={{fontSize:11,color:"#2a4050"}}>lalu isi di Excel, upload balik</span>
            </div>
            <div
              onDragOver={e=>{e.preventDefault();setDragging(true);}}
              onDragLeave={()=>setDragging(false)}
              onDrop={e=>{e.preventDefault();setDragging(false);!loading&&e.dataTransfer.files[0]&&handleFile(e.dataTransfer.files[0]);}}
              onClick={()=>!loading&&fileRef.current.click()}
              style={{...C.drop,borderColor:dragging?"#34d399":"rgba(52,211,153,0.2)",background:dragging?"rgba(52,211,153,0.05)":"rgba(255,255,255,0.02)", opacity:loading?0.5:1}}
            >
              {loading ? (
                <div style={{padding:20}}>
                  <div style={{fontSize:24, marginBottom:10}} className="spin">⏳</div>
                  <div style={{fontSize:14, fontWeight:700, color:"#34d399"}}>MEMPROSES DATA BESAR...</div>
                  <div style={{fontSize:10, color:"#6b7fa3", marginTop:5}}>Harap tunggu, sedang menyinkronkan 17.000+ baris data</div>
                </div>
              ) : (
                <>
                  <div style={{fontSize:32,marginBottom:8}}>📂</div>
                  <div style={{fontSize:13,color:"#6b7fa3",fontWeight:600}}>Klik atau drag file CSV / EXCEL di sini</div>
                  <div style={{fontSize:11,color:"#2a4050",marginTop:4}}>Format CSV atau Kertas Kerja 1721-A1 (.xlsx)</div>
                </>
              )}
              <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" style={{display:"none"}} onChange={e=>e.target.files[0]&&handleFile(e.target.files[0])}/>
            </div>
            {csvText && (
              <div style={{marginTop:10}}>
                <div style={{fontSize:11,color:"#34d399",marginBottom:6}}>✅ {csvInspection.rows.length} baris terdeteksi dari CSV</div>
                <div style={{maxHeight:120,overflowY:"auto",background:"rgba(0,0,0,0.2)",borderRadius:8,padding:"8px 12px",border:"1px solid rgba(255,255,255,0.06)"}}>
                  <pre style={{fontSize:10,color:"#3d5040",margin:0}}>{csvText.slice(0,400)}{csvText.length>400?"…":""}</pre>
                </div>
                <div style={{...C.infoBox, marginTop: 10}}>
                  <div style={{fontSize:10,fontWeight:700,color:"#6b7fa3",letterSpacing:"1.2px",marginBottom:8}}>PREVIEW MAPPING CSV</div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit, minmax(220px, 1fr))",gap:8}}>
                    {csvInspection.recognizedHeaders.map((item) => (
                      <div key={`${item.target}-${item.source}`} style={{fontSize:11,color:"#e8eaf0",background:"rgba(0,0,0,0.2)",padding:"8px 10px",borderRadius:8}}>
                        <strong style={{color:"#93c5fd"}}>{item.source}</strong> → {item.target}
                      </div>
                    ))}
                  </div>
                  {!!csvInspection.unmatchedHeaders.length && (
                    <div style={{marginTop:10,fontSize:11,color:"#fde68a",lineHeight:1.5}}>
                      Belum termapping: {csvInspection.unmatchedHeaders.join(", ")}
                    </div>
                  )}
                  {!!csvInspection.warnings?.length && (
                    <div style={{marginTop:10,fontSize:11,color:"#fca5a5",lineHeight:1.5}}>
                      Warning import: {csvInspection.warnings.slice(0, 5).join(" | ")}
                    </div>
                  )}
                  {!!missingRequiredMappings.length && (
                    <div style={{marginTop:10,fontSize:11,color:"#fca5a5",lineHeight:1.5}}>
                      Kolom wajib belum termapping: {missingRequiredMappings.join(", ")}
                    </div>
                  )}
                </div>
                <div style={{fontSize:10,color:"#6b7fa3",marginTop:8,lineHeight:1.5}}>
                  Jika kolom payroll Anda namanya berbeda, tambahkan dulu aliasnya di panel "Mapping Alias Komponen Payroll" di atas.
                </div>
                <button onClick={()=>setCsvText("")} style={{...C.ghostBtn,marginTop:8,fontSize:11}}>Ganti File</button>
              </div>
            )}
          </div>
        )}

        {/* ── HITUNG BUTTON ── */}
        <button onClick={hitung} disabled={loading || !importReady} style={{...C.calcBtn, opacity:loading||!importReady?0.7:1, cursor:loading?"wait":!importReady?"not-allowed":"pointer"}}>
          {progress !== null 
            ? `⏳ SEDANG MENGHITUNG... ${progress}%` 
            : `⚡ Hitung Semua (${sourceEmployees.length} karyawan)`}
        </button>

        {/* ── RESULTS ── */}
        {results && totals && (
          <div style={{marginTop:24}}>

            {/* Summary cards */}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",marginBottom:12, gap: 16, flexWrap: "wrap"}}>
              <div style={{fontSize:10,fontWeight:700,color:"#64748b",letterSpacing:"1.5px"}}>HASIL PERHITUNGAN DETAIL · {period}</div>
              <input 
                placeholder="Cari nama karyawan..." 
                value={search} 
                onChange={e=>setSearch(e.target.value)}
                style={{...C.configInput, minWidth:260, padding:"8px 14px", fontSize:12}}
              />
            </div>
            {taxPosition && (
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit, minmax(280px, 1fr))",gap:12,marginBottom:12}}>
                <div style={C.infoBox}>
                  <div style={{fontSize:10,fontWeight:700,color:"#6b7fa3",letterSpacing:"1.2px",marginBottom:8}}>
                    KEWAJIBAN VS SUDAH DISETOR
                  </div>
                  <div style={{display:"grid",gap:8}}>
                    {[
                      ["PPh Setahun", rpSigned(taxPosition.annualPph)],
                      ["PPh Jan-Nov / Sebelumnya", rpSigned(taxPosition.priorWithheld)],
                      ["PPh Masa Ini", rpSigned(taxPosition.currentObligation)],
                    ].map(([label, value]) => (
                      <div key={label} style={{display:"flex",justifyContent:"space-between",gap:12,fontSize:11}}>
                        <span style={{color:"#6b7fa3"}}>{label}</span>
                        <strong style={{color:"#e8eaf0"}}>{value}</strong>
                      </div>
                    ))}
                  </div>
                </div>
                <div style={C.infoBox}>
                  <div style={{fontSize:10,fontWeight:700,color:"#6b7fa3",letterSpacing:"1.2px",marginBottom:8}}>
                    BEBAN PEGAWAI VS PERUSAHAAN
                  </div>
                  <div style={{display:"grid",gap:8}}>
                    {[
                      ["Ditanggung Pegawai", burdenFmt(taxPosition.employeeBurden), "#fca5a5"],
                      ["Ditunjang Perusahaan", burdenFmt(taxPosition.companyBurden), "#86efac"],
                    ].map(([label, value, color]) => (
                      <div key={label} style={{display:"flex",justifyContent:"space-between",gap:12,fontSize:11}}>
                        <span style={{color:"#6b7fa3"}}>{label}</span>
                        <strong style={{color}}>{value}</strong>
                      </div>
                    ))}
                  </div>
                  {mode === "final" && (
                    <div style={{marginTop:8,fontSize:10,color:"#6b7fa3",lineHeight:1.5}}>
                      Nilai di panel ini menunjukkan besar beban ekonomis masing-masing pihak, bukan tanda plus/minus kewajiban masa terakhir.
                    </div>
                  )}
                </div>
                {mode === "final" && settlement && (
                  <div style={C.infoBox}>
                    <div style={{fontSize:10,fontWeight:700,color:"#6b7fa3",letterSpacing:"1.2px",marginBottom:8}}>
                      SETTLEMENT INTERNAL
                    </div>
                    <div style={{display:"grid",gap:8}}>
                      {[
                        ["Tagih / Potong ke Ybs", rp(settlement.employeePayable), "#fca5a5"],
                        ["Refund ke Ybs", rp(settlement.employeeRefund), "#86efac"],
                        ["Beban / Setor Perusahaan", rp(settlement.companyPayable), "#fde68a"],
                        ["Refund ke Perusahaan", rp(settlement.companyRefund), "#93c5fd"],
                      ].map(([label, value, color]) => (
                        <div key={label} style={{display:"flex",justifyContent:"space-between",gap:12,fontSize:11}}>
                          <span style={{color:"#6b7fa3"}}>{label}</span>
                          <strong style={{color}}>{value}</strong>
                        </div>
                      ))}
                    </div>
                    <div style={{marginTop:8,fontSize:10,color:"#6b7fa3",lineHeight:1.5}}>
                      Panel ini menunjukkan arah penyelesaian internal antara perusahaan dan ybs setelah posisi pajak masa terakhir dihitung.
                    </div>
                  </div>
                )}
              </div>
            )}
            {mode === "final" && settlement && settlementBreakdown && (
              <div style={{...C.infoBox, marginBottom: 12}}>
                <div style={{fontSize:10,fontWeight:700,color:"#6b7fa3",letterSpacing:"1.2px",marginBottom:8}}>
                  REKAP PENGEMBALIAN (REFUND)
                </div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit, minmax(320px, 1fr))",gap:10}}>
                  {[
                    ["Refund ke Pribadi / Ybs", settlement.employeeRefund, settlementBreakdown.employeeRefundCount, settlementBreakdown.employeeRefundRows, "#86efac"],
                    ["Refund ke Perusahaan", settlement.companyRefund, settlementBreakdown.companyRefundCount, settlementBreakdown.companyRefundRows, "#93c5fd"],
                  ].map(([title, totalValue, count, rows, color]) => (
                    <div key={title} style={{background:"rgba(0,0,0,0.2)",borderRadius:10,padding:10,border:"1px solid rgba(255,255,255,0.06)"}}>
                      <div style={{display:"flex",justifyContent:"space-between",gap:12,alignItems:"baseline",marginBottom:8}}>
                        <div style={{fontSize:11,fontWeight:700,color:"#e8eaf0"}}>{title}</div>
                        <div style={{fontSize:12,fontWeight:800,color}}>{rp(totalValue)}</div>
                      </div>
                      <div style={{fontSize:10,color:"#6b7fa3",marginBottom:8}}>
                        {count} pegawai
                        {search ? ` · hasil pencarian: ${rows.length}` : ""}
                      </div>
                      <div style={{maxHeight:180,overflowY:"auto",borderTop:"1px solid rgba(255,255,255,0.06)",paddingTop:8}}>
                        {(rows.length ? rows.slice(0, 30) : []).map((r) => (
                          <div key={`${title}-${r.id}`} style={{display:"grid",gridTemplateColumns:"1fr auto",gap:10,padding:"6px 0",borderBottom:"1px dashed rgba(255,255,255,0.05)"}}>
                            <div style={{fontSize:11,color:"#cbd5e1"}}>
                              <div style={{fontWeight:600,color:"#e8eaf0"}}>{r.nama}</div>
                              <div style={{fontSize:10,color:"#6b7fa3"}}>{r.status} · PPh Masa Ini {rpSigned(r.pphMasaIni)}</div>
                            </div>
                            <div style={{fontSize:11,fontWeight:700,color,textAlign:"right"}}>{rp(r.amount)}</div>
                          </div>
                        ))}
                        {!rows.length && (
                          <div style={{fontSize:10,color:"#6b7fa3",padding:"6px 0"}}>
                            Tidak ada data refund pada kategori ini.
                          </div>
                        )}
                        {rows.length > 30 && (
                          <div style={{fontSize:10,color:"#6b7fa3",paddingTop:6}}>
                            Menampilkan 30 baris teratas dari {rows.length} data.
                          </div>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {visibleComparison && comparison && comparison.compared > 0 && (
              <div style={{...C.infoBox, marginBottom: 12}}>
                <div style={{fontSize:10,fontWeight:700,color:"#6b7fa3",letterSpacing:"1.2px",marginBottom:8}}>
                  ANALISIS SELISIH VS KERTAS KERJA TIM
                </div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit, minmax(160px, 1fr))",gap:8,marginBottom:10}}>
                  {[
                    [`Dibandingkan`, `${comparison.compared} pegawai`, "#93c5fd"],
                    [`Mismatch > Rp100`, `${comparison.mismatches} pegawai`, "#fca5a5"],
                    [`App lebih besar`, `${comparison.appHigher} pegawai`, "#fde68a"],
                    [`App lebih kecil`, `${comparison.appLower} pegawai`, "#86efac"],
                    [`Total Selisih`, rpSigned(comparison.totalDiff), comparison.totalDiff >= 0 ? "#fca5a5" : "#86efac"],
                  ].map(([label, value, color]) => (
                    <div key={label} style={{background:"rgba(0,0,0,0.2)",borderRadius:8,padding:"10px 12px"}}>
                      <div style={{fontSize:9,color:"#6b7fa3",marginBottom:4}}>{label}</div>
                      <div style={{fontSize:12,fontWeight:800,color}}>{value}</div>
                    </div>
                  ))}
                </div>
                <div style={{fontSize:10,color:"#6b7fa3",marginBottom:8}}>
                  {search
                    ? `Hasil pencarian "${search}" pada daftar selisih terbesar: ${visibleComparison.topDiffs.length} baris.`
                    : "Daftar 10 selisih terbesar berdasarkan `|PPh App - PPh Excel|`."}
                </div>
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,minWidth:560}}>
                    <thead>
                      <tr style={{background:"rgba(255,255,255,0.03)"}}>
                        {["Nama","Status","App","Kertas Kerja","Selisih"].map((h)=>(
                          <th key={h} style={{...C.Th}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {visibleComparison.topDiffs.map((r) => (
                        <tr key={`cmp-${r.id}`} style={C.Tr}>
                          <td style={{...C.Td,fontWeight:600,color:"#e8eaf0"}}>{r.nama}</td>
                          <td style={C.Td}>{r.status}</td>
                          <td style={{...C.Td,textAlign:"right",color:"#fde68a"}}>{rpSigned(r.pphTotal)}</td>
                          <td style={{...C.Td,textAlign:"right",color:"#93c5fd"}}>{rpSigned(r.pphExcel)}</td>
                          <td style={{...C.Td,textAlign:"right",fontWeight:700,color:Math.abs(r.pphTotal-r.pphExcel)>100?"#f87171":"#34d399"}}>
                            {rpSigned(r.pphTotal-r.pphExcel)}
                          </td>
                        </tr>
                      ))}
                      {visibleComparison.topDiffs.length === 0 && (
                        <tr style={C.Tr}>
                          <td colSpan={5} style={{...C.Td,textAlign:"center",color:"#6b7fa3"}}>
                            Tidak ada nama yang cocok dengan pencarian di daftar selisih terbesar.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* EXPORT BAR */}
            <div style={C.exportBar}>
              <span style={{fontSize:11,color:"#3d4f6e",fontWeight:600}}>📤 Export:</span>
              <button onClick={doExport} disabled={exporting} style={{...C.exportBtn,"--c":"#86efac","--bg":"rgba(134,239,172,0.1)","--br":"rgba(134,239,172,0.25)",background:"rgba(134,239,172,0.1)",color:"#86efac",border:"1px solid rgba(134,239,172,0.25)",opacity:exporting?.6:1}}>
                {exporting?"⏳ Menyiapkan...":"📊 Download Excel (3 Sheet)"}
              </button>
              <button onClick={doExportPDF} disabled={exportingPDF} style={{...C.exportBtn,background:"rgba(147,197,253,0.08)",color:"#93c5fd",border:"1px solid rgba(147,197,253,0.2)", opacity: exportingPDF ? 0.6 : 1}}>
                {exportingPDF ? "⏳ Menyiapkan PDF..." : "📄 PDF Slip Massal"}
              </button>
              {mode === "final" && (
                <button
                  onClick={doExportSettlement}
                  disabled={exportingSettlement}
                  style={{...C.exportBtn,background:"rgba(134,239,172,0.08)",color:"#4ade80",border:"1px solid rgba(134,239,172,0.2)",opacity:exportingSettlement?0.6:1}}
                >
                  {exportingSettlement ? "⏳ Menyiapkan Settlement..." : "📌 Export Settlement Refund"}
                </button>
              )}
              <button onClick={()=>alert("🌟 Format e-SPT tersedia di versi Premium!\nRp 299.000/bulan")} style={{...C.exportBtn,background:"rgba(253,230,138,0.08)",color:"#fde68a",border:"1px solid rgba(253,230,138,0.2)"}}>
                🧾 Format e-SPT ⭐
              </button>
            </div>

            {/* TABLE */}
            <div style={{overflowX:"auto",marginTop:4}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,minWidth:700}}>
                <thead>
                  <tr style={{background:"rgba(255,255,255,0.03)"}}>
                    {[["#","",36],["nama","Nama",150],["status","Status",58],["bruto","Bruto",105],["pphK","PPh Karyawan",115],["pphP","PPh Perusahaan",120],["pphExcel","Versi Excel",115],["selisih","Selisih",110],["takehome","Take Home",108],["","",36]].map(([k,lbl,w])=>(
                      <th key={k+lbl} onClick={()=>k&&doSort(k)} style={{...C.Th,width:w,cursor:k?"pointer":"default"}} className={["status","bruto","pphExcel","selisih"].includes(k) ? "hide-mobile" : ""}>
                        {lbl}{sortK===k?(sortD===1?" ↑":" ↓"):""}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((r, i) => [
                    <tr key={r.id} onClick={() => setExpandId(expandId === r.id ? null : r.id)} style={{...C.Tr, cursor: "pointer", background:expandId===r.id?"rgba(52,211,153,0.12)":i%2?"transparent":"rgba(255,255,255,0.01)"}}>
                      <td style={{...C.Td,color:"#6b7fa3",textAlign:"center"}}>{i+1}</td>
                      <td style={{...C.Td,fontWeight:600,color:"#e8eaf0",maxWidth:150}}>{r.nama||"—"}</td>
                      <td style={{...C.Td,textAlign:"center"}} className="hide-mobile"><span style={{background:"rgba(255,255,255,0.06)",padding:"2px 6px",borderRadius:5,fontSize:10}}>{r.status}</span></td>
                      <td style={{...C.Td,textAlign:"right",color:"#cbd5e1"}} className="hide-mobile">{rp(r.bruto)}</td>
                      <td style={{...C.Td,textAlign:"right",color:"#fca5a5",fontWeight:600}}>{taxFlowFmt(r.pphK)}</td>
                      <td style={{...C.Td,textAlign:"right",color:"#86efac",fontWeight:600}}>{taxFlowFmt(r.pphP)}</td>
                      <td style={{...C.Td,textAlign:"right",color:"#93c5fd",fontWeight:600}} className="hide-mobile">{rpSigned(r.pphExcel)}</td>
                      <td style={{...C.Td,textAlign:"right",color:Math.abs(r.pphTotal - r.pphExcel) > 100 ? "#f87171" : "#34d399", fontWeight:700}} className="hide-mobile">{rpSigned(r.pphTotal - r.pphExcel)}</td>
                      <td style={{...C.Td,textAlign:"right",color:"#fde68a",fontWeight:700}}>{rp(r.takehome)}</td>
                      <td style={{...C.Td,textAlign:"center",color:"#3d4f6e"}}>{expandId === r.id ? "▲" : "▼"}</td>
                    </tr>,
                    expandId === r.id && (
                      <tr key={r.id + "-det"} style={{background: "rgba(52,211,153,0.05)"}}>
                        <td colSpan={9} style={{padding: "15px 30px", borderBottom: "1px solid rgba(52,211,153,0.2)"}}>
                          <div style={{display: "flex", gap: 40, alignItems: "center"}}>
                            <div style={{flex: 1}}>
                              <div style={{fontSize: 10, color: "#6b7fa3", marginBottom: 8, fontWeight: 700, letterSpacing: 1}}>ANALISIS PROPORSI BEBAN</div>
                              <div style={{display: "flex", height: 12, borderRadius: 6, overflow: "hidden", marginBottom: 10, background: "#1e293b"}}>
                                <div style={{width: `${r.rasio}%`, background: "#34d399", transition: "width 0.5s"}} title="Ditunjang Perusahaan" />
                                <div style={{width: `${100 - r.rasio}%`, background: "#f87171", transition: "width 0.5s"}} title="Dipotong Karyawan" />
                              </div>
                              <div style={{display: "flex", justifyContent: "space-between", fontSize: 11}}>
                                <div style={{color: "#34d399"}}>● Ditunjang: {r.rasio.toFixed(1)}% ({taxFlowFmt(r.pphP)})</div>
                                <div style={{color: "#f87171"}}>● Dipotong: {(100-r.rasio).toFixed(1)}% ({taxFlowFmt(r.pphK)})</div>
                              </div>
                            </div>
                            <div style={{flex: 1, borderLeft: "1px solid rgba(255,255,255,0.1)", paddingLeft: 30}}>
                              <div style={{fontSize: 10, color: "#6b7fa3", marginBottom: 8, fontWeight: 700, letterSpacing: 1}}>CASH FLOW REAL</div>
                              {r.isFinal && (
                                <div style={{marginBottom:12,fontSize:11,fontWeight:700,color:r.pphTotal<0?"#86efac":r.pphTotal>0?"#fca5a5":"#93c5fd"}}>
                                  {r.pphTotal < 0 ? "Lebih potong / potensi pengembalian ke pegawai" : r.pphTotal > 0 ? "Masih ada kewajiban masa terakhir" : "Tidak ada selisih masa terakhir"}
                                </div>
                              )}
                              <div style={{display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px 20px"}}>
                                <div>
                                  <div style={{fontSize: 9, color: "#475569"}}>Take Home Pay</div>
                                  <div style={{fontSize: 14, fontWeight: 800, color: "#fde68a"}}>{rp(r.takehome)}</div>
                                </div>
                                <div>
                                  <div style={{fontSize: 9, color: "#475569"}}>Beban Perusahaan</div>
                                  <div style={{fontSize: 14, fontWeight: 800, color: "#e8eaf0"}}>{rp(r.beban)}</div>
                                </div>
                              </div>
                            </div>
                          </div>
                          {r.isFinal && (
                            <div style={{display:"grid",gridTemplateColumns:"repeat(4, minmax(160px, 1fr))",gap:10,marginTop:16}}>
                              {[
                                ["Tagih ke Ybs", Math.max(0, r.pphK), "#fca5a5", "Potong / tagih tambahan ke pegawai"],
                                ["Refund ke Ybs", Math.max(0, -r.pphK), "#86efac", "Kembalikan ke pegawai"],
                                ["Beban Perusahaan", Math.max(0, r.pphP), "#fde68a", "Setor / tanggung oleh perusahaan"],
                                ["Refund ke Perusahaan", Math.max(0, -r.pphP), "#93c5fd", "Kembalikan / kompensasikan ke perusahaan"],
                              ].map(([label, value, color, note]) => (
                                <div key={label} style={{padding:"10px",background:"rgba(0,0,0,0.2)",borderRadius:10,border:"1px solid rgba(255,255,255,0.05)"}}>
                                  <div style={{fontSize:9,color:"#475569",marginBottom:4}}>{label}</div>
                                  <div style={{fontSize:13,fontWeight:800,color}}>{rp(value)}</div>
                                  <div style={{fontSize:9,color:"#6b7fa3",marginTop:4,lineHeight:1.4}}>{note}</div>
                                </div>
                              ))}
                            </div>
                          )}
                          <div style={{display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10, marginTop: 20}}>
                            {[
                              ["Gaji Pokok", r.gajiPokok, "money", r.policyUsed?.gajiPokok ? "✅" : "❌"],
                              ["Tunjangan", r.tunjangan, "money", r.policyUsed?.tunjangan ? "✅" : "❌"],
                              ["Lembur", r.lembur, "money", r.policyUsed?.lembur ? "✅" : "❌"],
                              ["Bonus/THR", r.bonus, "money", r.policyUsed?.bonus ? "✅" : "❌"],
                              ["BPJS", r.bpjs, "money", "➖"],
                              ["Tunjangan PPh", r.tunj, "signed", "🏢"],
                              [r.isAnnual || r.isFinal ? "PKP Setahun" : "Neto Basis TER", r.isAnnual || r.isFinal ? r.pkp : r.neto, "money", "📊"],
                              [r.isFinal ? "PPh 21 Masa Terakhir" : "PPh 21 Total", r.pphTotal, "signed", "🧾"],
                              ...(r.isFinal ? [["PPh Setahun", r.annualPphTotal, "signed", "📘"], ["PPh s.d. Lalu", r.pphSebelumnya, "signed", "📗"]] : []),
                            ].map(([l, v, fmt, icon]) => (
                              <div key={l} style={{padding: "10px", background: "rgba(0,0,0,0.2)", borderRadius: 10, border: "1px solid rgba(255,255,255,0.05)"}}>
                                <div style={{fontSize: 9, color: "#475569", marginBottom: 4}}>{icon} {l}</div>
                                <div style={{fontSize: 12, fontWeight: 700, color: "#e8eaf0"}}>{fmt === "signed" ? rpSigned(v) : rp(v)}</div>
                              </div>
                            ))}
                          </div>
                        </td>
                      </tr>
                    )
                  ])}
                </tbody>
                <tfoot>
                  <tr style={{background:"rgba(52,211,153,0.05)",borderTop:"2px solid rgba(52,211,153,0.15)"}}>
                    <td colSpan={3} style={{...C.Td,fontWeight:700,color:"#34d399",padding:"10px"}}>TOTAL ({results.length})</td>
                    <td style={{...C.Td,textAlign:"right",fontWeight:700,color:"#e8eaf0"}}>{rp(totals.bruto)}</td>
                    <td style={{...C.Td,textAlign:"right",fontWeight:700,color:"#fca5a5"}}>{taxFlowFmt(totals.pphK)}</td>
                    <td style={{...C.Td,textAlign:"right",fontWeight:700,color:"#86efac"}}>{taxFlowFmt(totals.pphP)}</td>
                    <td style={{...C.Td,textAlign:"center",color:"#3d4f6e"}}>—</td>
                    <td style={{...C.Td,textAlign:"center",color:"#3d4f6e"}}>—</td>
                    <td style={{...C.Td,textAlign:"right",fontWeight:700,color:"#fde68a"}}>{rp(totals.takehome)}</td>
                    <td/>
                  </tr>
                </tfoot>
              </table>
            </div>
            {results.length > 100 && !search && (
              <div style={{textAlign:"center",padding:10,fontSize:10,color:"#6b7fa3",background:"rgba(0,0,0,0.15)"}}>
                ... Menampilkan 100 baris pertama. Gunakan fitur "Cari" untuk melihat karyawan lain ...
              </div>
            )}
            <div style={{fontSize:10,color:"#1e2d3a",textAlign:"right",marginTop:6}}>Klik ▼ untuk detail · Klik header kolom untuk sort</div>
          </div>
        )}
      </div>
      <div style={{fontSize:10,color:"#111d2a",textAlign:"center",marginTop:24,paddingBottom:8}}>© 2026 Mamuyy PPh 21 · PMK 168/2023 · Alat bantu payroll internal · Bukan bukti potong resmi (1721-A1)</div>
    </div>
  );
}
