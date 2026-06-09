// src/app/(supplier)/supplier-dashboard/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import ExcelJS from "exceljs";
import { collection, getDocs, query, where, doc, updateDoc } from "firebase/firestore";
import { db, auth } from "@/lib/firebase"; // Imported auth natively for direct logout processing
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

  // Filter Modal Toggle & Input Parameter States
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

  // Export Progress Flag
  const [isExportingExcel, setIsExportingExcel] = useState(false);

  useEffect(() => {
    if (!loading && (!profile || profile.role !== "supplier")) {
      router.push("/login");
    }
  }, [profile, loading, router]);

  useEffect(() => {
    const supplierProfile = profile as any;
    if (!loading && profile?.role === "supplier" && supplierProfile?.supplierNo) {
      fetchSupplierWorkspaceData();
    }
  }, [profile, loading]);

  const fetchSupplierWorkspaceData = async () => {
    setIsDataLoading(true);
    const supplierProfile = profile as any;
    try {
      const materialsSnapshot = await getDocs(collection(db, "materials"));
      const matMap: Record<string, any> = {};
      materialsSnapshot.forEach((doc) => {
        matMap[doc.id] = doc.data();
      });
      setMaterialsMap(matMap);

      const q = query(
        collection(db, "rfq_routing"),
        where("supplierNo", "==", supplierProfile?.supplierNo || "")
      );
      const snapshot = await getDocs(q);
      const list: RFQItem[] = [];
      snapshot.forEach((doc) => {
        list.push({ id: doc.id, ...doc.data() } as RFQItem);
      });
      setRfqs(list);
    } catch (err) {
      console.error("Error pulling isolated vendor payload arrays:", err);
    } finally {
      setIsDataLoading(false);
    }
  };

  // HANDLES SECURE VENDOR LOGOUT ROUTING
  const handleSupplierLogout = async () => {
    try {
      await signOut(auth);
      router.push("/login");
    } catch (err) {
      console.error("Error signing out supplier account session:", err);
      alert("Failed to securely clear session tokens.");
    }
  };

  const startEditing = (item: RFQItem) => {
    setEditingId(item.id);
    setBidPrice(item.offeredPrice !== null ? item.offeredPrice.toString() : "");
    setLeadTime(item.leadTime || "");
    setVendorNotes(item.supplierNote || "");
  };

  const cancelEditing = () => {
    setEditingId(null);
    setBidPrice("");
    setLeadTime("");
    setVendorNotes("");
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
      fetchSupplierWorkspaceData(); 
    } catch (err) {
      console.error("Failed to commit supplier bid data:", err);
      alert("Error saving your bid.");
    } finally {
      setIsSaving(false);
    }
  };

  const handleExportTableToExcel = async () => {
    setIsExportingExcel(true);
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Quote Proposal");

      worksheet.views = [{ showGridLines: true }];

      worksheet.columns = [
        { key: "rfqId", width: 16 },
        { key: "itemNo", width: 14 },
        { key: "desc", width: 38 },
        { key: "qty", width: 10 },
        { key: "uom", width: 8 },
        { key: "buyer", width: 16 }, 
        { key: "price", width: 18 },
        { key: "total", width: 18 },
        { key: "leadTime", width: 16 },
        { key: "notes", width: 32 }
      ];

      worksheet.mergeCells("A2:C2");
      const titleCell = worksheet.getCell("A2");
      titleCell.value = (profile?.companyName || "VENDOR PROCUREMENT").toUpperCase();
      titleCell.font = { name: "Segoe UI", size: 16, bold: true, color: { argb: "1E3A8A" } };

      worksheet.mergeCells("A3:C3");
      const subtitleCell = worksheet.getCell("A3");
      subtitleCell.value = "Administrative Bidding Terminal Workspace Quote Proposal";
      subtitleCell.font = { name: "Segoe UI", size: 10, italic: true, color: { argb: "475569" } };

      worksheet.getCell("A5").value = "FROM (Supplier Profile):";
      worksheet.getCell("A5").font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "475569" } };
      
      worksheet.getCell("A6").value = "Supplier Code:";
      worksheet.getCell("A6").font = { name: "Segoe UI", size: 10, bold: true };
      worksheet.getCell("B6").value = (profile as any)?.supplierNo || "—";
      worksheet.getCell("B6").font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "1E3A8A" } };

      worksheet.getCell("A7").value = "Account Contact:";
      worksheet.getCell("A7").font = { name: "Segoe UI", size: 10, bold: true };
      worksheet.getCell("B7").value = (profile as any)?.contactName || (profile as any)?.name || "—";
      worksheet.getCell("B7").font = { name: "Segoe UI", size: 10 };

      worksheet.getCell("A8").value = "Email Destination:";
      worksheet.getCell("A8").font = { name: "Segoe UI", size: 10, bold: true };
      worksheet.getCell("B8").value = profile?.email || "—";
      worksheet.getCell("B8").font = { name: "Segoe UI", size: 10 };

      worksheet.getCell("G5").value = "TO (Procurement Destination):";
      worksheet.getCell("G5").font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "475569" } };

      worksheet.getCell("G6").value = "Austal USA";
      worksheet.getCell("G6").font = { name: "Segoe UI", size: 11, bold: true };
      worksheet.getCell("G7").value = "100 Austal Way";
      worksheet.getCell("G7").font = { name: "Segoe UI", size: 10 };
      worksheet.getCell("G8").value = "Mobile, AL 36602";
      worksheet.getCell("G8").font = { name: "Segoe UI", size: 10 };

      worksheet.getCell("G2").value = "DATE GENERATED:";
      worksheet.getCell("G2").font = { name: "Segoe UI", size: 10, bold: true };
      worksheet.getCell("H2").value = new Date().toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
      worksheet.getCell("H2").font = { name: "Segoe UI", size: 10 };

      const tableHeaders = ["RFQ ID", "Item #", "Material Description", "Qty", "UOM", "Buyer Assigned", "Offered Unit Price", "Extended Total", "Lead Time", "Supplier Notes"];
      const headerRowIndex = 11;
      const headerRow = worksheet.getRow(headerRowIndex);
      headerRow.height = 26;

      tableHeaders.forEach((text, idx) => {
        const cell = headerRow.getCell(idx + 1);
        cell.value = text;
        cell.font = { name: "Segoe UI", bold: true, color: { argb: "FFFFFF" }, size: 10 };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "1E3A8A" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin", color: { argb: "CBD5E1" } },
          left: { style: "thin", color: { argb: "CBD5E1" } },
          bottom: { style: "thin", color: { argb: "CBD5E1" } },
          right: { style: "thin", color: { argb: "CBD5E1" } }
        };
      });

      let currentRowIndex = 12;
      filteredRows.forEach((item, index) => {
        const matchingMaterial = materialsMap[item.materialId];
        const row = worksheet.getRow(currentRowIndex);
        row.height = 20;

        row.getCell(1).value = item.rfqId || "—";
        row.getCell(2).value = item.itemNumber || "—";
        row.getCell(3).value = item.description || "";
        row.getCell(4).value = Number(item.quantity || 0);
        row.getCell(5).value = item.uom || "EA";
        row.getCell(6).value = item.buyer || "—";
        row.getCell(7).value = item.offeredPrice !== null ? Number(item.offeredPrice) : 0;
        row.getCell(8).value = { formula: `=D${currentRowIndex}*G${currentRowIndex}`, result: 0 };
        row.getCell(9).value = item.leadTime || "—";
        row.getCell(10).value = item.supplierNote || "—";

        row.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
        row.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
        row.getCell(3).alignment = { horizontal: "left", vertical: "middle" };
        row.getCell(4).alignment = { horizontal: "right", vertical: "middle" };
        row.getCell(5).alignment = { horizontal: "center", vertical: "middle" };
        row.getCell(6).alignment = { horizontal: "left", vertical: "middle" };
        row.getCell(7).alignment = { horizontal: "right", vertical: "middle" };
        row.getCell(7).numFmt = "$#,##0.00";
        row.getCell(8).alignment = { horizontal: "right", vertical: "middle" };
        row.getCell(8).numFmt = "$#,##0.00";
        row.getCell(8).font = { name: "Segoe UI", bold: true, size: 10 };
        row.getCell(9).alignment = { horizontal: "left", vertical: "middle" };
        row.getCell(10).alignment = { horizontal: "left", vertical: "middle" };

        for (let colIdx = 1; colIdx <= 10; colIdx++) {
          const cell = row.getCell(colIdx);
          if (colIdx !== 8) cell.font = { name: "Segoe UI", size: 10 };
          cell.border = {
            top: { style: "thin", color: { argb: "CBD5E1" } },
            left: { style: "thin", color: { argb: "CBD5E1" } },
            bottom: { style: "thin", color: { argb: "CBD5E1" } },
            right: { style: "thin", color: { argb: "CBD5E1" } }
          };
          if (index % 2 === 1) {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "F8FAFC" } };
          }
        }
        currentRowIndex++;
      });

      const totalsRow = worksheet.getRow(currentRowIndex);
      totalsRow.height = 24;
      totalsRow.getCell(7).value = "Estimated Total:";
      totalsRow.getCell(7).font = { name: "Segoe UI", bold: true, size: 10 };
      totalsRow.getCell(7).alignment = { horizontal: "right", vertical: "middle" };

      const grandTotalCell = totalsRow.getCell(8);
      grandTotalCell.value = { formula: `=SUM(H12:H${currentRowIndex - 1})`, result: 0 };
      grandTotalCell.font = { name: "Segoe UI", bold: true, color: { argb: "1E3A8A" }, size: 11 };
      grandTotalCell.alignment = { horizontal: "right", vertical: "middle" };
      grandTotalCell.numFmt = "$#,##0.00";
      grandTotalCell.border = {
        top: { style: "thin", color: { argb: "000000" } },
        bottom: { style: "double", color: { argb: "000000" } }
      };

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `Quote_Proposal_${profile?.companyName || "Vendor"}_${new Date().toISOString().substring(0,10)}.xlsx`;
      anchor.click();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Excel generation tracking error:", err);
      alert("Error building quote proposal template workbook mapping records.");
    } finally {
      setIsExportingExcel(false);
    }
  };

  const formatTimestamp = (ts: any, mode: "dateOnly" | "fullTime") => {
    if (!ts) return <span className="text-slate-300">—</span>;
    const date = ts.toDate ? ts.toDate() : new Date(ts);
    if (mode === "dateOnly") {
      return <span className="font-medium text-slate-600 font-mono text-xs">{date.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })}</span>;
    }
    return (
      <span className="font-medium text-slate-600 block whitespace-nowrap text-xs">
        {date.toLocaleDateString("en-US", { month: "short", day: "numeric" })}
        <span className="text-[10px] text-slate-400 block font-normal">{date.toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit" })}</span>
      </span>
    );
  };

  const filteredRows = rfqs.filter((item) => {
    const computedRfqId = (item.rfqId || "").toLowerCase();
    const matchesRfqId = computedRfqId.includes(filterRfqId.trim().toLowerCase());
    const matchesItemNo = (item.itemNumber || "").toLowerCase().includes(filterItemNumber.trim().toLowerCase());
    const matchesDesc = (item.description || "").toLowerCase().includes(filterDescription.trim().toLowerCase());
    const matchesBuyer = (item.buyer || "").toLowerCase().includes(filterBuyer.trim().toLowerCase());
    return matchesRfqId && matchesItemNo && matchesDesc && matchesBuyer;
  });

  const clearFilterFields = () => {
    setFilterRfqId("");
    setFilterItemNumber("");
    setFilterDescription("");
    setFilterBuyer("");
  };

  if (loading) return <div className="p-8 text-sm text-slate-500">Verifying security parameters...</div>;

  const currentSupplierNo = (profile as any)?.supplierNo || "——";

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      
      <div className="p-4 bg-transparent rounded">
        
        {/* TOP LEVEL LOGISTICS ROUTING SHIPYARD METADATA HEADER BANNER */}
        <header className="mb-8 flex flex-col md:flex-row justify-between items-start md:items-end border-b border-slate-200 pb-5 gap-6">
          <div className="w-full md:w-auto flex flex-col sm:flex-row sm:items-start justify-between sm:gap-12 md:gap-0">
            <div>
              <h1 className="text-3xl font-extrabold tracking-tight text-slate-900">
                {profile?.companyName || "Vendor Procurement"}
              </h1>
              <h2 className="text-lg font-bold text-slate-600 mt-1">Bidding Terminal</h2>
              <div className="text-xs text-slate-500 mt-2 space-y-0.5">
                <p><span className="font-semibold text-slate-700">Supplier No:</span> {currentSupplierNo}</p>
                <p><span className="font-semibold text-slate-700">Account Contact:</span> {(profile as any)?.contactName || (profile as any)?.name || "Active Session User"}</p>
                <p><span className="font-semibold text-slate-700">Email Address:</span> {profile?.email || "—"}</p>
              </div>
            </div>
            
            {/* INLINE HEADER ACCESS CONTROL TERMINAL BUTTON */}
            <div className="mt-4 sm:mt-1">
              <button
                onClick={handleSupplierLogout}
                className="text-xs font-bold text-red-600 hover:text-red-800 bg-red-50 hover:bg-red-100 border border-red-200 px-3 py-1.5 rounded transition-all shadow-sm"
              >
                🚪 Secure Logout
              </button>
            </div>
          </div>
          
          {/* HARDCODED DESTINATION FOR PROPOSAL PACKETS SUBMITTALS ROUTING SHOWN GRAPHICALLY */}
          <div className="bg-white border border-slate-200 p-3.5 rounded-lg shadow-sm text-xs min-w-[210px] ml-auto md:ml-0">
            <h4 className="font-bold text-slate-400 uppercase tracking-wider mb-1">To:</h4>
            <div className="text-slate-800 font-medium space-y-0.5">
              <p className="font-bold text-sm text-slate-900">Austal USA</p>
              <p>100 Austal Way</p>
              <p>Mobile, AL 36602</p>
            </div>
          </div>
        </header>

        {/* UTILITY BAR OPERATIONS STRIP */}
        <div className="mb-4 flex flex-wrap items-center justify-end gap-2">
          <button
            onClick={() => setIsFilterModalOpen(true)}
            className="flex items-center text-sm font-semibold text-slate-700 bg-white border border-slate-300 px-3 py-1.5 rounded-md hover:bg-slate-50 shadow-sm transition-all"
          >
            🔍 Filter Queue { (filterRfqId || filterItemNumber || filterDescription || filterBuyer) && <span className="ml-1.5 h-2 w-2 rounded-full bg-blue-600" /> }
          </button>
          
          <button
            onClick={handleExportTableToExcel}
            disabled={isExportingExcel}
            className="text-sm font-bold text-blue-700 bg-blue-50 border border-blue-200 px-4 py-1.5 rounded-md hover:bg-blue-100 transition-colors shadow-sm disabled:opacity-50"
          >
            {isExportingExcel ? "Building Sheet Proposal..." : "📑 Generate Quote Proposal"}
          </button>
        </div>

        {/* WORKSPACE LIVE REQUIREMENT TRACKING TABLE CARD */}
        <div className="bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden">
          <div className="px-6 py-4 border-b border-slate-200 bg-slate-50/70">
            <h3 className="text-sm font-bold uppercase text-slate-700 tracking-wider">Open Material Items Request Log</h3>
          </div>

          <div className="overflow-x-auto">
            {isDataLoading ? (
              <div className="p-12 text-center text-slate-500">Loading open material parameters...</div>
            ) : (
              <table className="w-full text-left border-collapse text-sm">
                <thead className="bg-slate-100 text-slate-700 font-semibold text-xs border-b border-slate-200">
                  <tr>
                    <th className="py-3 px-4 text-center">RFQ ID</th>
                    <th className="py-3 px-6">Item #</th>
                    <th className="py-3 px-6">Description</th>
                    <th className="py-3 px-6 text-right">Qty</th>
                    <th className="py-3 px-6">UOM</th>
                    <th className="py-3 px-6 font-semibold text-slate-700">Buyer</th> 
                    <th className="py-3 px-6">Your Price ($)</th>
                    <th className="py-3 px-6">Lead Time</th>
                    <th className="py-3 px-6">Notes</th>
                    <th className="py-3 px-6">Date Uploaded</th>
                    <th className="py-3 px-6">Quote Date</th>
                    <th className="py-3 px-6 text-center">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 text-slate-800">
                  {filteredRows.length === 0 ? (
                    <tr>
                      <td colSpan={12} className="py-12 text-center text-slate-400">No material requests matched your criteria.</td>
                    </tr>
                  ) : (
                    filteredRows.map((item) => {
                      const isEditing = editingId === item.id;
                      const matchingMaterialDoc = materialsMap[item.materialId];
                      const rawUploadedTimestamp = matchingMaterialDoc?.timestamp || null;

                      return (
                        <tr key={item.id} className={`hover:bg-slate-50/50 transition-colors ${isEditing ? 'bg-blue-50/30' : ''}`}>
                          <td className="py-4 px-4 text-center font-mono font-bold text-xs text-slate-700 bg-slate-50/20">{item.rfqId || "—"}</td>
                          <td className="py-4 px-6 font-mono font-medium text-slate-900">{item.itemNumber}</td>
                          <td className="py-4 px-6 max-w-xs truncate" title={item.description}>{item.description}</td>
                          <td className="py-4 px-6 text-right font-medium">{item.quantity}</td>
                          <td className="py-4 px-6 text-slate-500">{item.uom}</td>
                          <td className="py-4 px-6 text-slate-600 font-medium whitespace-nowrap">{item.buyer || <span className="text-slate-300">—</span>}</td>

                          <td className="py-3 px-4">
                            {isEditing ? (
                              <input
                                type="number"
                                step="0.01"
                                value={bidPrice}
                                onChange={(e) => setBidPrice(e.target.value)}
                                className="w-24 rounded border border-slate-300 px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-blue-500 text-slate-900 font-bold"
                                placeholder="0.00"
                              />
                            ) : item.offeredPrice !== null ? (
                              <span className="font-semibold text-slate-900">${item.offeredPrice.toFixed(2)}</span>
                            ) : (
                              <span className="text-slate-300 font-medium">Pending Entry</span>
                            )}
                          </td>

                          <td className="py-3 px-4">
                            {isEditing ? (
                              <input
                                type="text"
                                value={leadTime}
                                onChange={(e) => setLeadTime(e.target.value)}
                                className="w-28 rounded border border-slate-300 px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-blue-500 text-slate-900 font-medium"
                                placeholder="e.g. 2 weeks"
                              />
                            ) : (
                              <span className="text-slate-700">{item.leadTime || <span className="text-slate-300">—</span>}</span>
                            )}
                          </td>

                          <td className="py-3 px-4">
                            {isEditing ? (
                              <input
                                type="text"
                                value={vendorNotes}
                                onChange={(e) => setVendorNotes(e.target.value)}
                                className="w-full min-w-[150px] rounded border border-slate-300 px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-blue-500 text-slate-900"
                                placeholder="Add optional notes..."
                              />
                            ) : (
                              <span className="text-xs text-slate-500 max-w-xs truncate block" title={item.supplierNote}>
                                {item.supplierNote || <span className="text-slate-300">—</span>}
                              </span>
                            )}
                          </td>

                          <td className="py-3 px-6 whitespace-nowrap">{formatTimestamp(rawUploadedTimestamp, "dateOnly")}</td>
                          <td className="py-3 px-6 text-xs">{formatTimestamp(item.timestamp, "fullTime")}</td>

                          <td className="py-3 px-6 text-center whitespace-nowrap">
                            {isEditing ? (
                              <div className="flex items-center justify-center gap-2">
                                <button onClick={() => handleSaveBid(item.id)} disabled={isSaving} className="rounded bg-emerald-600 px-2.5 py-1 text-xs font-semibold text-white hover:bg-emerald-500">{isSaving ? "Saving..." : "Save"}</button>
                                <button onClick={cancelEditing} className="rounded border border-slate-300 bg-white px-2.5 py-1 text-xs text-slate-700 hover:bg-slate-50">Cancel</button>
                              </div>
                            ) : (
                              <button onClick={() => startEditing(item)} className="inline-flex items-center rounded bg-blue-50 px-3 py-1.5 text-xs font-semibold text-blue-700 hover:bg-blue-100">
                                {item.offeredPrice !== null ? "Edit Bid" : "Quote Price"}
                              </button>
                            )}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            )}
          </div>
        </div>
      </div>

      {/* FILTER PARAMETERS OVERLAY MODAL */}
      {isFilterModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="w-full max-w-sm rounded-lg bg-white p-6 shadow-xl border border-slate-200">
            <div className="border-b border-slate-200 pb-3 mb-4 flex justify-between items-center">
              <h3 className="text-md font-bold text-slate-900">Filter Procurement Items</h3>
              <button type="button" onClick={clearFilterFields} className="text-xs text-blue-600 hover:text-blue-800 font-semibold">Reset All</button>
            </div>
            <div className="space-y-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">RFQ ID Reference</label>
                <input 
                  type="text" 
                  value={filterRfqId} 
                  onChange={(e) => setFilterRfqId(e.target.value)} 
                  className="w-full text-sm rounded border border-slate-300 px-3 py-2 uppercase font-mono text-slate-900 placeholder:text-slate-300 focus:outline-none focus:ring-1 focus:ring-blue-500" 
                  placeholder="e.g. PROJECT-1" 
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Item # Identifier</label>
                <input 
                  type="text" 
                  value={filterItemNumber} 
                  onChange={(e) => setFilterItemNumber(e.target.value)} 
                  className="w-full text-sm rounded border border-slate-300 px-3 py-2 font-mono text-slate-900 placeholder:text-slate-300 focus:outline-none focus:ring-1 focus:ring-blue-500" 
                  placeholder="e.g. 1001-A" 
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Material Description Keyword</label>
                <input 
                  type="text" 
                  value={filterDescription} 
                  onChange={(e) => setFilterDescription(e.target.value)} 
                  className="w-full text-sm rounded border border-slate-300 px-3 py-2 text-slate-900 placeholder:text-slate-300 focus:outline-none focus:ring-1 focus:ring-blue-500" 
                  placeholder="e.g. Steel Pipe" 
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Buyer</label>
                <input 
                  type="text" 
                  value={filterBuyer} 
                  onChange={(e) => setFilterBuyer(e.target.value)} 
                  className="w-full text-sm rounded border border-slate-300 px-3 py-2 text-slate-900 placeholder:text-slate-300 focus:outline-none focus:ring-1 focus:ring-blue-500" 
                  placeholder="e.g. James Rush" 
                />
              </div>
            </div>
            <div className="flex justify-end gap-2 border-t border-slate-200 pt-4 mt-6">
              <button type="button" onClick={() => setIsFilterModalOpen(false)} className="w-full rounded bg-blue-600 py-2 text-center text-sm font-semibold text-white hover:bg-blue-500">Apply Active Parameters ({filteredRows.length} Rows)</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}