// src/app/(supplier)/supplier-dashboard/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import ExcelJS from "exceljs";
import { collection, onSnapshot, query, where, doc, updateDoc, getDocs } from "firebase/firestore";
import { db, auth } from "@/lib/firebase"; 
import { useAuth } from "@/context/AuthContext";
import { useRouter } from "next/navigation";
import { signOut } from "firebase/auth";

interface RFQItem {
  id: string;
  materialId: string;
  rfqId: string; 
  itemNumber: string;
  description: string;
  quantity: number;
  uom: string;
  buyer: string; 
  status: "Pending" | "Completed";
  isHot?: boolean;
  offeredPrice: number | null;
  leadTime: string | null;
  supplierNote: string;
  timestamp: any; 
}

export default function SupplierDashboard() {
  const { profile, loading } = useAuth();
  const router = useRouter();

  const [rfqs, setRfqs] = useState<RFQItem[]>([]);
  const [materialsMap, setMaterialsMap] = useState<Record<string, any>>({});
  const [isDataLoading, setIsDataLoading] = useState(false);

  // Filter Modal Parameter States
  const [isFilterModalOpen, setIsFilterModalOpen] = useState(false);
  const [filterRfqId, setFilterRfqId] = useState("");
  const [filterItemNumber, setFilterItemNumber] = useState("");
  const [filterDescription, setFilterDescription] = useState("");
  const [filterBuyer, setFilterBuyer] = useState("");

  // Inline Bidding Entry States
  const [editingId, setEditingId] = useState<string | null>(null);
  const [bidPrice, setBidPrice] = useState<string>("");
  const [leadTime, setLeadTime] = useState<string>("");
  const [vendorNotes, setVendorNotes] = useState<string>("");
  const [isSaving, setIsSaving] = useState(false);

  const [isExportingExcel, setIsExportingExcel] = useState(false);
  const [isGeneratingProposal, setIsGeneratingProposal] = useState(false);

  useEffect(() => {
    if (!loading && (!profile || profile.role !== "supplier")) {
      router.push("/login");
    }
  }, [profile, loading, router]);

  useEffect(() => {
    const supplierProfile = profile as any;
    if (loading || profile?.role !== "supplier" || !supplierProfile?.supplierNo) return;

    setIsDataLoading(true);

    getDocs(collection(db, "materials")).then((materialsSnapshot) => {
      const matMap: Record<string, any> = {};
      materialsSnapshot.forEach((mDoc) => {
        matMap[mDoc.id] = mDoc.data();
      });
      setMaterialsMap(matMap);
    }).catch(err => console.error("Error setting up lookup maps:", err));

    const routingQuery = query(
      collection(db, "rfq_routing"),
      where("supplierNo", "==", supplierProfile.supplierNo)
    );

    const unsubscribeRouting = onSnapshot(routingQuery, (snapshot) => {
      const list: RFQItem[] = [];
      snapshot.forEach((doc) => {
        list.push({ id: doc.id, ...doc.data() } as RFQItem);
      });
      setRfqs(list);
      setIsDataLoading(false);
    }, (err) => {
      console.error(err);
      setIsDataLoading(false);
    });

    return () => unsubscribeRouting();
  }, [profile, loading]);

  const handleSupplierLogout = async () => {
    try {
      await signOut(auth);
      router.push("/login");
    } catch (err) {
      console.error(err);
    }
  };

  const startEditing = (item: RFQItem) => {
    setEditingId(item.id);
    setBidPrice(item.offeredPrice !== null ? item.offeredPrice.toString() : "");
    setLeadTime(item.leadTime || "");
    setVendorNotes(item.supplierNote || "");
  };

  const handleSaveBid = async (rfqId: string) => {
    const parsedPrice = parseFloat(bidPrice);
    if (isNaN(parsedPrice) || parsedPrice < 0) {
      alert("Please enter a valid numeric pricing value.");
      return;
    }

    setIsSaving(true);
    try {
      const rfqDocRef = doc(db, "rfq_routing", rfqId);
      await updateDoc(rfqDocRef, {
        offeredPrice: parsedPrice,
        leadTime: leadTime.trim(),
        supplierNote: vendorNotes.trim(),
        status: "Completed",
        timestamp: new Date()
      });
      setEditingId(null);
    } catch (err) {
      console.error(err);
    } finally {
      setIsSaving(false);
    }
  };

  const handleExportRawExcel = async () => {
    setIsExportingExcel(true);
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Bids Workspace Data");
      worksheet.views = [{ showGridLines: true }];
      worksheet.columns = [
        { header: "RFQ ID", key: "rfqId", width: 16 },
        { header: "Item #", key: "itemNo", width: 14 },
        { header: "Description", key: "desc", width: 36 },
        { header: "Quantity", key: "qty", width: 10 },
        { header: "UOM", key: "uom", width: 8 },
        { header: "Buyer Assigned", key: "buyer", width: 16 }, 
        { header: "Your Offered Price ($)", key: "price", width: 20 },
        { header: "Lead Time", key: "leadTime", width: 16 }
      ];

      worksheet.getRow(1).height = 26;
      worksheet.getRow(1).font = { name: "Segoe UI", bold: true, color: { argb: "FFFFFF" } };
      worksheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "334155" } };
      worksheet.getRow(1).alignment = { horizontal: "center", vertical: "middle" };

      filteredRows.forEach((item, index) => {
        const row = worksheet.addRow({
          rfqId: `${item.isHot ? '🔥 ' : ''}${item.rfqId || '—'}`,
          itemNo: item.itemNumber || "—",
          desc: item.description || "",
          qty: Number(item.quantity || 0),
          uom: item.uom || "EA",
          buyer: item.buyer || "—",
          price: item.offeredPrice !== null ? Number(item.offeredPrice) : 0,
          leadTime: item.leadTime || "—"
        });
        row.eachCell((cell) => {
          cell.font = { name: "Segoe UI", size: 10 };
          if (index % 2 === 1) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "F8FAFC" } };
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `Supplier_Data_Export_${new Date().toISOString().substring(0,10)}.xlsx`;
      anchor.click();
    } catch (err) {
      console.error(err);
    } finally {
      setIsExportingExcel(false);
    }
  };

  const handleExportTableToExcel = async () => {
    setIsGeneratingProposal(true);
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Quote Proposal");
      worksheet.views = [{ showGridLines: true }];
      worksheet.columns = [
        { key: "rfqId", width: 16 }, { key: "itemNo", width: 14 }, { key: "desc", width: 38 },
        { key: "qty", width: 10 }, { key: "uom", width: 8 }, { key: "buyer", width: 16 }, 
        { key: "price", width: 18 }, { key: "total", width: 18 }, { key: "leadTime", width: 16 }, { key: "notes", width: 32 }
      ];

      worksheet.mergeCells("A2:C2");
      worksheet.getCell("A2").value = (profile?.companyName || "VENDOR PROCUREMENT").toUpperCase();
      worksheet.getCell("A2").font = { name: "Segoe UI", size: 16, bold: true, color: { argb: "1E3A8A" } };

      worksheet.getCell("G6").value = "Austal USA";
      worksheet.getCell("G6").font = { name: "Segoe UI", size: 11, bold: true };
      worksheet.getCell("G7").value = "100 Austal Way";
      worksheet.getCell("G8").value = "Mobile, AL 36602";

      const tableHeaders = ["RFQ ID", "Item #", "Material Description", "Qty", "UOM", "Buyer Assigned", "Offered Unit Price", "Extended Total", "Lead Time", "Supplier Notes"];
      const headerRow = worksheet.getRow(11);
      headerRow.height = 26;
      tableHeaders.forEach((text, idx) => {
        const cell = headerRow.getCell(idx + 1);
        cell.value = text;
        cell.font = { name: "Segoe UI", bold: true, color: { argb: "FFFFFF" }, size: 10 };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "1E3A8A" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
      });

      let currentRowIndex = 12;
      filteredRows.forEach((item, index) => {
        const row = worksheet.getRow(currentRowIndex);
        row.height = 20;
        row.getCell(1).value = `${item.isHot ? '🔥 ' : ''}${item.rfqId || '—'}`;
        row.getCell(2).value = item.itemNumber || "—";
        row.getCell(3).value = item.description || "";
        row.getCell(4).value = Number(item.quantity || 0);
        row.getCell(5).value = item.uom || "EA";
        row.getCell(6).value = item.buyer || "—";
        row.getCell(7).value = item.offeredPrice !== null ? Number(item.offeredPrice) : 0;
        row.getCell(8).value = { formula: `=D${currentRowIndex}*G${currentRowIndex}` };
        row.getCell(9).value = item.leadTime || "—";
        row.getCell(10).value = item.supplierNote || "—";

        for (let colIdx = 1; colIdx <= 10; colIdx++) {
          const cell = row.getCell(colIdx);
          cell.alignment = { vertical: "middle", horizontal: colIdx === 4 || colIdx === 7 || colIdx === 8 ? "right" : (colIdx === 3 || colIdx === 6 || colIdx === 9 || colIdx === 10 ? "left" : "center") };
          if (colIdx === 7 || colIdx === 8) cell.numFmt = "$#,##0.00";
          cell.border = { top: { style: "thin", color: { argb: "CBD5E1" } }, left: { style: "thin", color: { argb: "CBD5E1" } }, bottom: { style: "thin", color: { argb: "CBD5E1" } }, right: { style: "thin", color: { argb: "CBD5E1" } } };
          if (index % 2 === 1) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "F8FAFC" } };
        }
        currentRowIndex++;
      });

      const totalsRow = worksheet.getRow(currentRowIndex);
      totalsRow.getCell(7).value = "Estimated Total:";
      totalsRow.getCell(7).font = { name: "Segoe UI", bold: true };
      totalsRow.getCell(8).value = { formula: `=SUM(H12:H${currentRowIndex - 1})` };
      totalsRow.getCell(8).font = { name: "Segoe UI", bold: true, color: { argb: "1E3A8A" } };
      totalsRow.getCell(8).numFmt = "$#,##0.00";

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `Quote_Proposal_${profile?.companyName || "Vendor"}_${new Date().toISOString().substring(0,10)}.xlsx`;
      anchor.click();
    } catch (err) {
      console.error(err);
    } finally {
      setIsGeneratingProposal(false);
    }
  };

  const filteredRows = rfqs.filter((item) => {
    return (
      (item.rfqId || "").toLowerCase().includes(filterRfqId.trim().toLowerCase()) &&
      (item.itemNumber || "").toLowerCase().includes(filterItemNumber.trim().toLowerCase()) &&
      (item.description || "").toLowerCase().includes(filterDescription.trim().toLowerCase()) &&
      (item.buyer || "").toLowerCase().includes(filterBuyer.trim().toLowerCase())
    );
  });

  const currentSupplierNo = (profile as any)?.supplierNo || "——";

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <div className="p-4 bg-transparent rounded">
        <header className="mb-8 flex flex-col md:flex-row justify-between items-start md:items-end border-b border-slate-200 pb-5 gap-6">
          <div>
            <h1 className="text-3xl font-extrabold tracking-tight text-slate-900">{profile?.companyName || "Vendor Procurement"}</h1>
            <h2 className="text-lg font-bold text-slate-600 mt-1">Bidding Terminal</h2>
            <div className="text-xs text-slate-500 mt-2 space-y-0.5">
              <p><span className="font-semibold text-slate-700">Supplier No:</span> {currentSupplierNo}</p>
              <p><span className="font-semibold text-slate-700">Account Contact:</span> {(profile as any)?.contactName || (profile as any)?.name || "Active User"}</p>
            </div>
          </div>
          <div className="bg-white border border-slate-200 p-3.5 rounded-lg shadow-sm text-xs min-w-[210px]">
            <h4 className="font-bold text-slate-400 uppercase tracking-wider mb-1">To:</h4>
            <div className="text-slate-800 font-medium space-y-0.5">
              <p className="font-bold text-sm text-slate-900">Austal USA</p>
              <p>100 Austal Way</p>
              <p>Mobile, AL 36602</p>
            </div>
          </div>
        </header>

        <div className="mb-4 flex flex-wrap items-center justify-between gap-2 bg-slate-100/60 p-2.5 rounded-lg border border-slate-200/80">
          <button type="button" onClick={() => setIsFilterModalOpen(true)} className="flex items-center text-sm font-semibold text-slate-700 bg-white border border-slate-300 px-3 py-1.5 rounded-md hover:bg-slate-50 shadow-sm">🔍 Filter Queue</button>
          <div className="flex flex-wrap items-center gap-2">
            <button type="button" onClick={handleExportRawExcel} disabled={isExportingExcel} className="text-sm font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-3 py-1.5 rounded-md hover:bg-emerald-100 shadow-sm">{isExportingExcel ? "Exporting..." : "📊 Export to Excel"}</button>
            <button type="button" onClick={handleExportTableToExcel} disabled={isGeneratingProposal} className="text-sm font-bold text-blue-700 bg-blue-50 border border-blue-200 px-4 py-1.5 rounded-md hover:bg-blue-100 shadow-sm">{isGeneratingProposal ? "Compiling..." : "📑 Generate Quote Proposal"}</button>
            <button type="button" onClick={handleSupplierLogout} className="text-sm font-bold text-red-600 hover:text-red-800 bg-white border border-red-200 px-3 py-1.5 rounded-md shadow-sm">🚪 Secure Logout</button>
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-sm">
              <thead className="bg-slate-100 text-slate-700 font-semibold text-xs border-b border-slate-200">
                <tr>
                  <th className="py-3 px-4 text-center">RFQ ID</th>
                  <th className="py-3 px-6">Item #</th>
                  <th className="py-3 px-6">Description</th>
                  <th className="py-3 px-6 text-right">Qty</th>
                  <th className="py-3 px-6">UOM</th>
                  <th className="py-3 px-6">Buyer</th> 
                  <th className="py-3 px-6">Your Price ($)</th>
                  <th className="py-3 px-6">Lead Time</th>
                  <th className="py-3 px-6">Notes</th>
                  <th className="py-3 px-6 text-center">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 text-slate-800">
                {filteredRows.map((item) => {
                  const isEditing = editingId === item.id;
                  return (
                    <tr key={item.id} className={`hover:bg-slate-50/50 transition-colors ${isEditing ? 'bg-blue-50/30' : ''} ${item.isHot ? 'bg-red-50/10 font-medium' : ''}`}>
                      <td className="py-4 px-4 text-center font-mono font-bold text-xs text-slate-700 bg-slate-50/20">
                        {/* FLAME EMOJI PREFIX CONDITIONAL LAYOUT LAYER INJECTION */}
                        <span className="flex items-center justify-center gap-1">
                          {item.isHot && <span title="High Priority Hot Requirement" className="text-red-500 animate-bounce">🔥</span>}
                          {item.rfqId || "—"}
                        </span>
                      </td>
                      <td className="py-4 px-6 font-mono text-slate-900">{item.itemNumber}</td>
                      <td className="py-4 px-6 max-w-xs truncate" title={item.description}>{item.description}</td>
                      <td className="py-4 px-6 text-right font-medium">{item.quantity}</td>
                      <td className="py-4 px-6 text-slate-500">{item.uom}</td>
                      <td className="py-4 px-6 text-slate-600 font-medium whitespace-nowrap">{item.buyer || "—"}</td>
                      <td className="py-3 px-4">
                        {isEditing ? <input type="number" step="0.01" value={bidPrice} onChange={(e) => setBidPrice(e.target.value)} className="w-24 rounded border border-slate-300 px-2 py-1 text-slate-900 font-bold" /> : (item.offeredPrice !== null ? <span className="font-semibold text-slate-900">${item.offeredPrice.toFixed(2)}</span> : <span className="text-slate-300">Pending</span>)}
                      </td>
                      <td className="py-3 px-4">
                        {isEditing ? <input type="text" value={leadTime} onChange={(e) => setLeadTime(e.target.value)} className="w-28 rounded border border-slate-300 px-2 py-1 text-slate-900" /> : item.leadTime || "—"}
                      </td>
                      <td className="py-3 px-4">
                        {isEditing ? <input type="text" value={vendorNotes} onChange={(e) => setVendorNotes(e.target.value)} className="w-full min-w-[150px] rounded border border-slate-300 px-2 py-1 text-slate-900" /> : <span className="text-xs text-slate-500 max-w-xs truncate block">{item.supplierNote || "—"}</span>}
                      </td>
                      <td className="py-3 px-6 text-center whitespace-nowrap">
                        {isEditing ? (
                          <div className="flex gap-2 justify-center"><button type="button" onClick={() => handleSaveBid(item.id)} className="rounded bg-emerald-600 px-2.5 py-1 text-xs font-bold text-white">Save</button><button type="button" onClick={() => setEditingId(null)} className="rounded border border-slate-300 bg-white px-2.5 py-1 text-xs">Cancel</button></div>
                        ) : <button type="button" onClick={() => startEditing(item)} className="rounded bg-blue-50 px-3 py-1.5 text-xs font-semibold text-blue-700 hover:bg-blue-100">Quote Price</button>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}