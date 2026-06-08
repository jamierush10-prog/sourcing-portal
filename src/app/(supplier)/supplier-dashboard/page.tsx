// src/app/(supplier)/supplier-dashboard/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import ExcelJS from "exceljs";
import { collection, getDocs, query, where, doc, updateDoc } from "firebase/firestore";
import { db } from "@/lib/firebase";
import { useAuth } from "@/context/AuthContext";
import { useRouter } from "next/navigation";

interface RFQItem {
  id: string;
  materialId: string;
  itemNumber: string;
  description: string;
  quantity: number;
  uom: string;
  status: "Pending" | "Completed";
  offeredPrice: number | null;
  leadTime: string | null;
  supplierNote: string;
  timestamp: any; // Submittal date stamp token
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

  // Inline Bidding Entry States
  const [editingId, setEditingId] = useState<string | null>(null);
  const [bidPrice, setBidPrice] = useState<string>("");
  const [leadTime, setLeadTime] = useState<string>("");
  const [vendorNotes, setVendorNotes] = useState<string>("");
  const [isSaving, setIsSaving] = useState(false);

  // Spreadsheet generation state
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
      // 1. Pull master material logs to capture creation dates natively
      const materialsSnapshot = await getDocs(collection(db, "materials"));
      const matMap: Record<string, any> = {};
      materialsSnapshot.forEach((doc) => {
        matMap[doc.id] = doc.data();
      });
      setMaterialsMap(matMap);

      // 2. Query routed items assigned to this vendor
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

  // EXCEL TABLE SPREADSHEET BUILDER TRIGGER
  const handleExportTableToExcel = async () => {
    setIsExportingExcel(true);
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("My Bids Workspace");

      worksheet.columns = [
        { header: "RFQ ID", key: "rfqId", width: 15 },
        { header: "Item #", key: "itemNo", width: 14 },
        { header: "Description", key: "desc", width: 36 },
        { header: "Quantity", key: "qty", width: 10 },
        { header: "UOM", key: "uom", width: 8 },
        { header: "Your Price ($)", key: "price", width: 16 },
        { header: "Lead Time", key: "leadTime", width: 16 },
        { header: "Notes", key: "notes", width: 30 },
        { header: "Date Uploaded", key: "dateUploaded", width: 16 },
        { header: "Quote Date Submitted", key: "quoteDate", width: 22 }
      ];

      worksheet.getRow(1).height = 26;
      worksheet.getRow(1).font = { name: "Segoe UI", bold: true, color: { argb: "FFFFFF" } };
      worksheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "1E3A8A" } }; // Dark blue theme
      worksheet.getRow(1).alignment = { horizontal: "center", vertical: "middle" };

      filteredRows.forEach((item) => {
        const matchingMaterial = materialsMap[item.materialId];
        const uploadedDateStr = matchingMaterial?.timestamp 
          ? (matchingMaterial.timestamp.toDate ? matchingMaterial.timestamp.toDate() : new Date(matchingMaterial.timestamp)).toLocaleDateString()
          : "—";

        const quoteDateStr = item.timestamp
          ? (item.timestamp.toDate ? item.timestamp.toDate() : new Date(item.timestamp)).toLocaleString()
          : "—";

        const row = worksheet.addRow({
          rfqId: item.materialId ? `RFQ-${item.materialId.substring(0, 5).toUpperCase()}` : "—",
          itemNo: item.itemNumber || "",
          desc: item.description || "",
          qty: Number(item.quantity || 0),
          uom: item.uom || "EA",
          price: item.offeredPrice !== null ? Number(item.offeredPrice) : "Pending",
          leadTime: item.leadTime || "—",
          notes: item.supplierNote || "",
          dateUploaded: uploadedDateStr,
          quoteDate: item.status === "Completed" ? quoteDateStr : "—"
        });

        row.height = 20;
        row.getCell("rfqId").alignment = { horizontal: "center", vertical: "middle" };
        row.getCell("itemNo").alignment = { horizontal: "center", vertical: "middle" };
        row.getCell("qty").alignment = { horizontal: "right", vertical: "middle" };
        row.getCell("uom").alignment = { horizontal: "center", vertical: "middle" };
        if (item.offeredPrice !== null) {
          row.getCell("price").numberFormat = "$#,##0.00";
          row.getCell("price").alignment = { horizontal: "right", vertical: "middle" };
        } else {
          row.getCell("price").alignment = { horizontal: "center", vertical: "middle" };
        }
      });

      worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell) => {
          cell.border = {
            top: { style: "thin", color: { argb: "CBD5E1" } },
            left: { style: "thin", color: { argb: "CBD5E1" } },
            bottom: { style: "thin", color: { argb: "CBD5E1" } },
            right: { style: "thin", color: { argb: "CBD5E1" } }
          };
          if (rowNumber > 1) {
            cell.font = { name: "Segoe UI", size: 10 };
            if (rowNumber % 2 === 0) {
              cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "F8FAFC" } };
            }
          }
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `Supplier_Bidding_Log_${new Date().toISOString().substring(0,10)}.xlsx`;
      anchor.click();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Spreadsheet save error:", err);
    } finally {
      setIsExportingExcel(false);
    }
  };

  // DATA TIMESTAMPS FORMATTING UTILITIES
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

  // MULTI-PARAMETER FILTER MATCHING LOOP LOGIC
  const filteredRows = rfqs.filter((item) => {
    const computedRfqId = item.materialId ? `RFQ-${item.materialId.substring(0, 5).toUpperCase()}` : "";
    const matchesRfqId = computedRfqId.toLowerCase().includes(filterRfqId.trim().toLowerCase());
    const matchesItemNo = (item.itemNumber || "").toLowerCase().includes(filterItemNumber.trim().toLowerCase());
    const matchesDesc = (item.description || "").toLowerCase().includes(filterDescription.trim().toLowerCase());
    return matchesRfqId && matchesItemNo && matchesDesc;
  });

  const clearFilterFields = () => {
    setFilterRfqId("");
    setFilterItemNumber("");
    setFilterDescription("");
  };

  if (loading) return <div className="p-8 text-sm text-slate-500">Verifying security parameters...</div>;

  const currentSupplierNo = (profile as any)?.supplierNo || "——";

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      {/* 1. DYNAMIC RESTRUCTURED HEADER BLOCKS TIMELINE */}
      <header className="mb-8 flex flex-col md:flex-row justify-between items-start md:items-end border-b border-slate-200 pb-5 gap-4">
        <div>
          {/* Main Title Header: Supplier Name Only */}
          <h1 className="text-3xl font-extrabold tracking-tight text-slate-900">
            {profile?.companyName || "Vendor"}
          </h1>
          {/* Sub Header: Bidding Terminal Text */}
          <h2 className="text-lg font-bold text-slate-600 mt-1">
            Bidding Terminal
          </h2>
          {/* Tertiary Line: Supplier Code Stamp */}
          <p className="text-xs text-slate-400 mt-1 font-medium">
            Supplier Code: <span className="font-mono font-bold text-blue-600 bg-blue-50 px-2 py-0.5 rounded ml-1">{currentSupplierNo}</span>
          </p>
        </div>
        
        {/* ACTION UTILITY MANAGEMENT RUNTIME CONTROLS BAR */}
        <div className="flex items-center gap-2 w-full md:w-auto">
          <button
            onClick={() => setIsFilterModalOpen(true)}
            className="flex items-center text-sm font-semibold text-slate-700 bg-white border border-slate-300 px-3 py-1.5 rounded-md hover:bg-slate-50 shadow-sm transition-all"
          >
            🔍 Filter Queue { (filterRfqId || filterItemNumber || filterDescription) && <span className="ml-1.5 h-2 w-2 rounded-full bg-blue-600" /> }
          </button>
          
          <button
            onClick={handleExportTableToExcel}
            disabled={isExportingExcel}
            className="text-sm font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-3 py-1.5 rounded-md hover:bg-emerald-100 transition-colors shadow-sm disabled:opacity-50"
          >
            {isExportingExcel ? "Generating Spreadsheet..." : "📊 Export Table to Excel"}
          </button>
        </div>
      </header>

      {/* MATERIALS ACTIONS MASTER QUEUE DATA GRID CONTAINER */}
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
                    <td colSpan={11} className="py-12 text-center text-slate-400">
                      No material requests matched your filtering criteria.
                    </td>
                  </tr>
                ) : (
                  filteredRows.map((item) => {
                    const isEditing = editingId === item.id;
                    const computedRfqId = item.materialId ? `RFQ-${item.materialId.substring(0, 5).toUpperCase()}` : "—";
                    
                    // Pull date uploaded from materials mapping object
                    const matchingMaterialDoc = materialsMap[item.materialId];
                    const rawUploadedTimestamp = matchingMaterialDoc?.timestamp || null;

                    return (
                      <tr key={item.id} className={`hover:bg-slate-50/50 transition-colors ${isEditing ? 'bg-blue-50/30' : ''}`}>
                        
                        {/* 1. RFQ ID COLUMN */}
                        <td className="py-4 px-4 text-center font-mono font-bold text-xs text-slate-400 bg-slate-50/20">
                          {computedRfqId}
                        </td>

                        {/* 2. ITEM NUMBER COLUMN */}
                        <td className="py-4 px-6 font-mono font-medium text-slate-900">{item.itemNumber}</td>
                        <td className="py-4 px-6 max-w-xs truncate" title={item.description}>{item.description}</td>
                        <td className="py-4 px-6 text-right font-medium">{item.quantity}</td>
                        <td className="py-4 px-6 text-slate-500">{item.uom}</td>
                        
                        {/* Unit Bid Pricing */}
                        <td className="py-3 px-4">
                          {isEditing ? (
                            <input
                              type="number"
                              step="0.01"
                              value={bidPrice}
                              onChange={(e) => setBidPrice(e.target.value)}
                              className="w-24 rounded border border-slate-300 px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-blue-500"
                              placeholder="0.00"
                            />
                          ) : item.offeredPrice !== null ? (
                            <span className="font-semibold text-slate-900">${item.offeredPrice.toFixed(2)}</span>
                          ) : (
                            <span className="text-slate-300 font-medium">Pending Entry</span>
                          )}
                        </td>

                        {/* Lead Time */}
                        <td className="py-3 px-4">
                          {isEditing ? (
                            <input
                              type="text"
                              value={leadTime}
                              onChange={(e) => setLeadTime(e.target.value)}
                              className="w-28 rounded border border-slate-300 px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-blue-500"
                              placeholder="e.g. 2 weeks"
                            />
                          ) : (
                            <span className="text-slate-700">{item.leadTime || <span className="text-slate-300">—</span>}</span>
                          )}
                        </td>

                        {/* Supplier Notes */}
                        <td className="py-3 px-4">
                          {isEditing ? (
                            <input
                              type="text"
                              value={vendorNotes}
                              onChange={(e) => setVendorNotes(e.target.value)}
                              className="w-full min-w-[150px] rounded border border-slate-300 px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-blue-500"
                              placeholder="Add optional notes..."
                            />
                          ) : (
                            <span className="text-xs text-slate-500 max-w-xs truncate block" title={item.supplierNote}>
                              {item.supplierNote || <span className="text-slate-300">—</span>}
                            </span>
                          )}
                        </td>

                        {/* DATE LINE ITEM WAS UPLOADED ADDED TO APP */}
                        <td className="py-3 px-6 whitespace-nowrap">
                          {formatTimestamp(rawUploadedTimestamp, "dateOnly")}
                        </td>

                        {/* Bid Proposal Timestamp Submission date */}
                        <td className="py-3 px-6 text-xs">
                          {formatTimestamp(item.timestamp, "fullTime")}
                        </td>

                        {/* Table Loop Actions panel */}
                        <td className="py-3 px-6 text-center whitespace-nowrap">
                          {isEditing ? (
                            <div className="flex items-center justify-center gap-2">
                              <button
                                onClick={() => handleSaveBid(item.id)}
                                disabled={isSaving}
                                className="rounded bg-emerald-600 px-2.5 py-1 text-xs font-semibold text-white shadow-sm hover:bg-emerald-500"
                              >
                                {isSaving ? "Saving..." : "Save"}
                              </button>
                              <button onClick={cancelEditing} className="rounded border border-slate-300 bg-white px-2.5 py-1 text-xs text-slate-700 hover:bg-slate-50">
                                Cancel
                              </button>
                            </div>
                          ) : (
                            <button
                              onClick={() => startEditing(item)}
                              className="inline-flex items-center rounded bg-blue-50 px-3 py-1.5 text-xs font-semibold text-blue-700 hover:bg-blue-100"
                            >
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

      {/* SLIDEOUT INTERACTIVE MULTI-PARAMETER FILTER MODAL OVERLAY */}
      {isFilterModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="w-full max-w-sm rounded-lg bg-white p-6 shadow-xl border border-slate-200">
            <div className="border-b border-slate-200 pb-3 mb-4 flex justify-between items-center">
              <h3 className="text-md font-bold text-slate-900">Filter Procurement Items</h3>
              <button 
                type="button" 
                onClick={clearFilterFields} 
                className="text-xs text-blue-600 hover:text-blue-800 font-semibold"
              >
                Reset All
              </button>
            </div>
            
            <div className="space-y-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">RFQ ID Reference</label>
                <input
                  type="text"
                  value={filterRfqId}
                  onChange={(e) => setFilterRfqId(e.target.value)}
                  className="w-full text-sm rounded border border-slate-300 px-3 py-2 focus:outline-none focus:ring-1 focus:ring-blue-500 uppercase font-mono"
                  placeholder="e.g. RFQs-A12B"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Item # Identifier</label>
                <input
                  type="text"
                  value={filterItemNumber}
                  onChange={(e) => setFilterItemNumber(e.target.value)}
                  className="w-full text-sm rounded border border-slate-300 px-3 py-2 focus:outline-none focus:ring-1 focus:ring-blue-500 font-mono"
                  placeholder="e.g. 1001-A"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Material Description Keyword</label>
                <input
                  type="text"
                  value={filterDescription}
                  onChange={(e) => setFilterDescription(e.target.value)}
                  className="w-full text-sm rounded border border-slate-300 px-3 py-2 focus:outline-none focus:ring-1 focus:ring-blue-500"
                  placeholder="e.g. Steel Pipe"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2