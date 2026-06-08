// src/app/(admin)/dashboard/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import Papa from "papaparse";
import ExcelJS from "exceljs";
import { collection, writeBatch, doc, getDocs, query, orderBy, where, updateDoc, deleteDoc } from "firebase/firestore";
import { db } from "@/lib/firebase";
import { useAuth } from "@/context/AuthContext";
import { useRouter } from "next/navigation";

interface MaterialItem {
  id: string;
  rfqId: string; 
  itemNumber: string;
  description: string;
  quantity: number;
  uom: string;
  status: "Pending" | "Sourced" | "Completed";
  quoteCount?: number;
}

interface SupplierProfile {
  id: string;
  supplierNo: string;
  companyName: string;
  contactName: string;
  email: string;
}

interface BidResponse {
  id: string;
  materialId: string;
  rfqId: string;
  itemNumber: string;
  description: string;
  supplierNo: string;
  offeredPrice: number | null;
  leadTime: string | null;
  supplierNote: string;
  status: "Pending" | "Completed";
  timestamp: any;
}

export default function AdminDashboard() {
  const { profile, loading } = useAuth();
  const router = useRouter();
  
  const [materials, setMaterials] = useState<MaterialItem[]>([]);
  const [suppliers, setSuppliers] = useState<SupplierProfile[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [feedbackMessage, setFeedbackMessage] = useState("");

  // Filter Modal Toggle & Input Parameter States
  const [isFilterModalOpen, setIsFilterModalOpen] = useState(false);
  const [filterRfqId, setFilterRfqId] = useState("");
  const [filterItemNumber, setFilterItemNumber] = useState("");
  const [filterDescription, setFilterDescription] = useState("");

  // CRUD Inline Editing States
  const [editingItemId, setEditingItemId] = useState<string | null>(null);
  const [editRfqId, setEditRfqId] = useState("");
  const [editItemNumber, setEditItemNumber] = useState("");
  const [editDescription, setEditDescription] = useState("");
  const [editQuantity, setEditQuantity] = useState<number>(0);
  const [editUom, setEditUom] = useState("");
  const [isSavingCrud, setIsSavingCrud] = useState(false);

  // Sourcing Modal States
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [selectedItem, setSelectedItem] = useState<MaterialItem | null>(null);
  const [selectedSupplierNos, setSelectedSupplierNos] = useState<string[]>([]);
  const [isRouting, setIsRouting] = useState(false);

  // Quotes Viewer Modal States
  const [isQuotesModalOpen, setIsQuotesModalOpen] = useState(false);
  const [releasedVendors, setReleasedVendors] = useState<BidResponse[]>([]);
  const [receivedQuotes, setReceivedQuotes] = useState<BidResponse[]>([]);
  const [isQuotesLoading, setIsQuotesLoading] = useState(false);

  // Excel Export Progress Spinner
  const [isExportingExcel, setIsExportingExcel] = useState(false);

  useEffect(() => {
    if (!loading && (!profile || profile.role !== "admin")) {
      router.push("/login");
    }
  }, [profile, loading, router]);

  useEffect(() => {
    if (profile?.role === "admin") {
      fetchMaterialsAndCounts();
      fetchSuppliers();
    }
  }, [profile]);

  const fetchMaterialsAndCounts = async () => {
    try {
      const materialsQuery = query(collection(db, "materials"), orderBy("itemNumber", "asc"));
      const materialsSnapshot = await getDocs(materialsQuery);
      const itemsList: MaterialItem[] = [];
      
      materialsSnapshot.forEach((doc) => {
        itemsList.push({ id: doc.id, ...doc.data(), quoteCount: 0 } as MaterialItem);
      });

      // Count only completed replies for the main grid badge metric
      const routingQuery = query(collection(db, "rfq_routing"), where("status", "==", "Completed"));
      const routingSnapshot = await getDocs(routingQuery);
      
      routingSnapshot.forEach((routingDoc) => {
        const data = routingDoc.data();
        const matchingMaterial = itemsList.find(item => item.id === data.materialId);
        if (matchingMaterial && matchingMaterial.quoteCount !== undefined) {
          matchingMaterial.quoteCount += 1;
        }
      });

      setMaterials(itemsList);
    } catch (err) {
      console.error("Error loading requirements matrix: ", err);
    }
  };

  const fetchSuppliers = async () => {
    try {
      const q = query(collection(db, "suppliers"), orderBy("companyName", "asc"));
      const snapshot = await getDocs(q);
      const list: SupplierProfile[] = [];
      snapshot.forEach((doc) => {
        list.push({ id: doc.id, ...doc.data() } as SupplierProfile);
      });
      setSuppliers(list);
    } catch (err) {
      console.error("Error mapping authorized vendors:", err);
    }
  };

  const handleCSVUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsUploading(true);
    setFeedbackMessage("Parsing structural file fields...");

    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: async (results) => {
        try {
          const batch = writeBatch(db);
          
          results.data.forEach((row: any) => {
            const newDocRef = doc(collection(db, "materials"));
            batch.set(newDocRef, {
              rfqId: row["RFQ ID"] || row["rfqId"] || "REQ-UNASSIGNED", 
              itemNumber: row["Item"] || row["itemNumber"] || "UNKNOWN",
              description: row["Description"] || row["description"] || "",
              quantity: Number(row["Qty"] || row["quantity"] || 0),
              uom: row["UOM"] || row["uom"] || "EA",
              status: "Pending",
              timestamp: new Date()
            });
          });

          await batch.commit();
          setFeedbackMessage("Requirements queue initialized.");
          fetchMaterialsAndCounts();
        } catch (error) {
          console.error("CSV upload batch failed:", error);
          setFeedbackMessage("Failed to process data columns.");
        } finally {
          setIsUploading(false);
        }
      },
    });
  };

  const downloadCsvTemplate = () => {
    const headers = ["RFQ ID", "Item", "Description", "Qty", "UOM"];
    const sampleRows = [
      ["REQ-2026-01", "1001-A", "3/4\" Structural Carbon Steel Plate A36", "12", "EA"],
      ["PROJECT-BLUE", "1002-B", "Grade 60 #5 Rebar Deformed Steel 20ft", "250", "LF"]
    ];
    
    const csvContent = [headers, ...sampleRows]
      .map(row => row.map(val => `"${val.replace(/"/g, '""')}"`).join(","))
      .join("\n");

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.setAttribute("download", "Material_Requirements_Template.csv");
    link.click();
    URL.revokeObjectURL(url);
  };

  // CRUD Inline Editing States
  const startEditingRow = (item: MaterialItem) => {
    setEditingItemId(item.id);
    setEditRfqId(item.rfqId || "");
    setEditItemNumber(item.itemNumber);
    setEditDescription(item.description);
    setEditQuantity(item.quantity);
    setEditUom(item.uom);
  };

  const cancelEditingRow = () => {
    setEditingItemId(null);
  };

  const handleUpdateItemRow = async (id: string) => {
    setIsSavingCrud(true);
    try {
      const docRef = doc(db, "materials", id);
      await updateDoc(docRef, {
        rfqId: editRfqId.trim(), 
        itemNumber: editItemNumber.trim(),
        description: editDescription.trim(),
        quantity: Number(editQuantity),
        uom: editUom.trim(),
      });
      setEditingItemId(null);
      fetchMaterialsAndCounts();
    } catch (err) {
      console.error("Failed to update item row:", err);
      alert("Error saving your edits.");
    } finally {
      setIsSavingCrud(false);
    }
  };

  const handleDeleteItemRow = async (id: string) => {
    if (!confirm("Are you sure you want to completely delete this line requirement?")) return;
    try {
      await deleteDoc(doc(db, "materials", id));
      fetchMaterialsAndCounts();
    } catch (err) {
      console.error("Failed to clear material document profile:", err);
    }
  };

  const openSourcingModal = (item: MaterialItem) => {
    setSelectedItem(item);
    setSelectedSupplierNos([]);
    setIsModalOpen(true);
  };

  const handleToggleSupplierSelection = (supplierNo: string) => {
    setSelectedSupplierNos((prev) =>
      prev.includes(supplierNo) ? prev.filter((no) => no !== supplierNo) : [...prev, supplierNo]
    );
  };

  const handleDispatchRFQ = async () => {
    if (!selectedItem || selectedSupplierNos.length === 0) return;
    setIsRouting(true);

    try {
      const batch = writeBatch(db);

      selectedSupplierNos.forEach((supNo) => {
        const rfqDocRef = doc(collection(db, "rfq_routing"));
        batch.set(rfqDocRef, {
          materialId: selectedItem.id,
          rfqId: selectedItem.rfqId || "—", 
          itemNumber: selectedItem.itemNumber,
          description: selectedItem.description,
          quantity: selectedItem.quantity,
          uom: selectedItem.uom,
          supplierNo: supNo,
          status: "Pending",
          offeredPrice: null,
          leadTime: null,
          supplierNote: "",
          timestamp: null, 
        });
      });

      const materialDocRef = doc(db, "materials", selectedItem.id);
      batch.update(materialDocRef, { status: "Sourced" });

      await batch.commit();
      setIsModalOpen(false);
      fetchMaterialsAndCounts();
    } catch (err) {
      console.error("RFQ routing dispatch failed:", err);
    } finally {
      setIsRouting(false);
    }
  };

  // AUDIT VIEWER: LOADS ALL ROUTED LOGS AND SEPARATES PENDING VS RECEIVED
  const openQuotesViewerModal = async (item: MaterialItem) => {
    setSelectedItem(item);
    setReleasedVendors([]);
    setReceivedQuotes([]);
    setIsQuotesLoading(true);
    setIsQuotesModalOpen(true);

    try {
      const q = query(collection(db, "rfq_routing"), where("materialId", "==", item.id));
      const snapshot = await getDocs(q);
      
      const pendingRelease: BidResponse[] = [];
      const completedBids: BidResponse[] = [];
      
      snapshot.forEach((doc) => {
        const data = doc.data() as BidResponse;
        if (data.status === "Completed") {
          completedBids.push({ id: doc.id, ...data });
        } else {
          pendingRelease.push({ id: doc.id, ...data });
        }
      });

      // Sort bids by price lowest to highest
      completedBids.sort((a, b) => (a.offeredPrice || 0) - (b.offeredPrice || 0));
      
      setReleasedVendors(pendingRelease);
      setReceivedQuotes(completedBids);
    } catch (err) {
      console.error("Failed compiling internal bid data elements:", err);
    } finally {
      setIsQuotesLoading(false);
    }
  };

  const handleExportAllQuotesToExcel = async () => {
    setIsExportingExcel(true);
    try {
      const q = query(collection(db, "rfq_routing"), where("status", "==", "Completed"));
      const snapshot = await getDocs(q);
      const allQuotes: any[] = [];
      
      snapshot.forEach((doc) => {
        allQuotes.push(doc.data());
      });

      const usersSnapshot = await getDocs(collection(db, "users"));
      const usersList: any[] = [];
      usersSnapshot.forEach((doc) => {
        usersList.push({ uid: doc.id, ...doc.data() });
      });

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Master Received Quotes Log");

      worksheet.columns = [
        { header: "RFQ ID", key: "rfqId", width: 16 },
        { header: "Item #", key: "itemNo", width: 14 },
        { header: "Material Description", key: "desc", width: 36 },
        { header: "Quantity", key: "qty", width: 10 },
        { header: "UOM", key: "uom", width: 8 },
        { header: "Supplier Corporate Name", key: "supplierName", width: 26 },
        { header: "Supplier Code", key: "supplierNo", width: 14 },
        { header: "Submitted By (User Email)", key: "userEmail", width: 28 },
        { header: "Quoted Unit Price ($)", key: "price", width: 20 },
        { header: "Lead Time Execution", key: "leadTime", width: 16 },
        { header: "Vendor Notes Summary", key: "notes", width: 40 },
        { header: "Quote Submittal Date Stamp", key: "dateStamp", width: 24 }
      ];

      worksheet.getRow(1).height = 26;
      worksheet.getRow(1).font = { name: "Segoe UI", bold: true, color: { argb: "FFFFFF" }, size: 11 };
      worksheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "0F172A" } }; 
      worksheet.getRow(1).alignment = { horizontal: "center", vertical: "middle" };

      allQuotes.forEach((quote) => {
        const vendorProfile = suppliers.find(s => s.supplierNo === quote.supplierNo);
        const userMatch = usersList.find(u => u.supplierNo === quote.supplierNo && u.role === "supplier");

        const parsedDate = quote.timestamp?.toDate ? quote.timestamp.toDate() : new Date(quote.timestamp);
        const formattedDate = parsedDate.toLocaleDateString("en-US", {
          month: "short", day: "numeric", year: "numeric", hour: "2-digit", minute: "2-digit"
        });

        const row = worksheet.addRow({
          rfqId: quote.rfqId || "—", 
          itemNo: quote.itemNumber || "—",
          desc: quote.description || "",
          qty: Number(quote.quantity || 0),
          uom: quote.uom || "EA",
          supplierName: vendorProfile ? vendorProfile.companyName : `Vendor Code: ${quote.supplierNo}`,
          supplierNo: quote.supplierNo || "",
          userEmail: userMatch ? userMatch.email : (vendorProfile ? vendorProfile.email : "—"),
          price: Number(quote.offeredPrice || 0),
          leadTime: quote.leadTime || "—",
          notes: quote.supplierNote || "",
          dateStamp: formattedDate
        });

        row.height = 20;
        row.getCell("rfqId").alignment = { horizontal: "center", vertical: "middle" };
        row.getCell("itemNo").alignment = { horizontal: "center", vertical: "middle" };
        row.getCell("qty").alignment = { horizontal: "right", vertical: "middle" };
        row.getCell("uom").alignment = { horizontal: "center", vertical: "middle" };
        row.getCell("supplierNo").alignment = { horizontal: "center", vertical: "middle" };
        row.getCell("price").alignment = { horizontal: "right", vertical: "middle" };
        row.getCell("price").numFmt = "$#,##0.00";
        row.getCell("dateStamp").alignment = { horizontal: "left", vertical: "middle" };
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
      anchor.download = `Material_Procurement_Master_Quotes_${new Date().toISOString().substring(0,10)}.xlsx`;
      anchor.click();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Excel generation execution error:", err);
    } finally {
      setIsExportingExcel(false);
    }
  };

  const filteredMaterials = materials.filter((item) => {
    const rfqField = (item.rfqId || "").toLowerCase();
    const itemNoField = (item.itemNumber || "").toLowerCase();
    const descField = (item.description || "").toLowerCase();

    return (
      rfqField.includes(filterRfqId.trim().toLowerCase()) &&
      itemNoField.includes(filterItemNumber.trim().toLowerCase()) &&
      descField.includes(filterDescription.trim().toLowerCase())
    );
  });

  const clearFilterFields = () => {
    setFilterRfqId("");
    setFilterItemNumber("");
    setFilterDescription("");
  };

  const formatTimestamp = (ts: any) => {
    if (!ts) return "—";
    const date = ts.toDate ? ts.toDate() : new Date(ts);
    return date.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric", hour: "2-digit", minute: "2-digit" });
  };

  const getSupplierName = (supNo: string) => {
    const match = suppliers.find(s => s.supplierNo === supNo);
    return match ? match.companyName : `Vendor (${supNo})`;
  };

  if (loading) return <div className="p-8 text-sm text-slate-500">Verifying credentials...</div>;

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <header className="mb-8 flex flex-col md:flex-row justify-between items-start md:items-center border-b border-slate-200 pb-4 gap-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">Material Procurement Console</h1>
          <p className="text-sm text-slate-500">Parse master material parameters and manage vendor distribution metrics</p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <button
            onClick={() => setIsFilterModalOpen(true)}
            className="text-sm font-semibold text-slate-700 bg-white border border-slate-300 px-3 py-1.5 rounded-md hover:bg-slate-50 shadow-sm transition-all"
          >
            🔍 Filter Queue { (filterRfqId || filterItemNumber || filterDescription) && <span className="ml-1.5 h-2 w-2 rounded-full bg-blue-600" /> }
          </button>
          <button
            onClick={handleExportAllQuotesToExcel}
            disabled={isExportingExcel}
            className="text-sm font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-3 py-1.5 rounded-md hover:bg-emerald-100 transition-colors shadow-sm disabled:opacity-50"
          >
            {isExportingExcel ? "Generating Spreadsheet..." : "📊 Export All Quotes to Excel"}
          </button>
          <button onClick={() => router.push("/users")} className="text-sm font-semibold text-purple-700 hover:text-purple-900 bg-purple-50 border border-purple-200 px-3 py-1.5 rounded-md transition-all">👤 User Accounts</button>
          <button onClick={() => router.push("/suppliers")} className="text-sm font-semibold text-blue-600 hover:text-blue-800 bg-blue-50 border border-blue-200 px-3 py-1.5 rounded-md transition-all">🏢 Suppliers Directory</button>
        </div>
      </header>

      {/* CSV Import Layout */}
      <div className="mb-8 p-4 bg-white border border-slate-200 rounded-lg shadow-sm flex flex-col sm:flex-row gap-4 justify-between items-start sm:items-center">
        <div>
          <h4 className="text-sm font-bold text-slate-800">Import Master Material Requirements</h4>
          <p className="text-xs text-slate-500">Upload bulk CSV listings or use the template to key in your custom RFQ/Requisition numbers</p>
        </div>
        <div className="flex flex-wrap items-center gap-3 w-full sm:w-auto justify-end">
          <button
            type="button"
            onClick={downloadCsvTemplate}
            className="text-xs font-bold text-blue-700 hover:text-blue-900 bg-blue-50 border border-blue-200 px-3 py-1.5 rounded transition-colors"
          >
            📥 Download Import Template
          </button>
          <input type="file" accept=".csv" onChange={handleCSVUpload} disabled={isUploading} className="text-xs text-slate-500 file:mr-2 file:py-1 file:px-3 file:rounded file:border-0 file:text-xs file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 cursor-pointer" />
          {feedbackMessage && <span className="text-xs font-medium text-slate-600 bg-slate-100 px-2 py-1 rounded">{feedbackMessage}</span>}
        </div>
      </div>

      {/* Master Grid Table Block */}
      <div className="bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-200 bg-slate-50/70">
          <h3 className="text-sm font-bold uppercase text-slate-700 tracking-wider">Master Requirements Queue</h3>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-sm">
            <thead className="bg-slate-100 text-slate-700 font-semibold text-xs border-b border-slate-200">
              <tr>
                <th className="py-3 px-4 text-center">RFQ ID</th>
                <th className="py-3 px-6">Item #</th>
                <th className="py-3 px-6">Description</th>
                <th className="py-3 px-6 text-right">Qty</th>
                <th className="py-3 px-6">UOM</th>
                <th className="py-3 px-6 text-center">Bids Received</th>
                <th className="py-3 px-6 text-center">Status</th>
                <th className="py-3 px-6 text-center">Console Operations</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 text-slate-800">
              {filteredMaterials.length === 0 ? (
                <tr><td colSpan={8} className="py-8 text-center text-slate-400">No matching requirements located.</td></tr>
              ) : (
                filteredMaterials.map((item) => {
                  const isEditingRow = editingItemId === item.id;

                  return (
                    <tr key={item.id} className={`hover:bg-slate-50/50 transition-colors ${isEditingRow ? 'bg-amber-50/40' : ''}`}>
                      <td className="py-4 px-4 text-center font-mono font-bold text-xs text-slate-700 bg-slate-50/20">
                        {isEditingRow ? (
                          <input type="text" value={editRfqId} onChange={(e) => setEditRfqId(e.target.value)} className="w-24 text-center text-xs rounded border border-slate-300 px-1 py-1 focus:outline-none focus:ring-1 focus:ring-blue-500 font-mono uppercase font-bold" placeholder="e.g. PROJECT-1" />
                        ) : (
                          item.rfqId || <span className="text-slate-300">—</span>
                        )}
                      </td>
                      <td className="py-4 px-6 font-mono font-medium text-slate-900">
                        {isEditingRow ? (
                          <input type="text" value={editItemNumber} onChange={(e) => setEditItemNumber(e.target.value)} className="w-28 text-sm rounded border border-slate-300 px-2 py-1 font-mono font-medium" />
                        ) : (
                          item.itemNumber
                        )}
                      </td>
                      <td className="py-4 px-6 max-w-xs truncate" title={item.description}>
                        {isEditingRow ? (
                          <input type="text" value={editDescription} onChange={(e) => setEditDescription(e.target.value)} className="w-full min-w-[180px] text-sm rounded border border-slate-300 px-2 py-1" />
                        ) : (
                          item.description
                        )}
                      </td>
                      <td className="py-4 px-6 text-right font-semibold">
                        {isEditingRow ? (
                          <input type="number" value={editQuantity} onChange={(e) => setEditQuantity(Number(e.target.value))} className="w-16 text-sm text-right rounded border border-slate-300 px-2 py-1 font-semibold" />
                        ) : (
                          item.quantity
                        )}
                      </td>
                      <td className="py-4 px-6 text-slate-500">
                        {isEditingRow ? (
                          <input type="text" value={editUom} onChange={(e) => setEditUom(e.target.value)} className="w-14 text-sm rounded border border-slate-300 px-2 py-1" />
                        ) : (
                          item.uom
                        )}
                      </td>
                      <td className="py-4 px-6 text-center">
                        <span className={`inline-flex items-center rounded-md px-2.5 py-0.5 text-xs font-bold ${
                          (item.quoteCount || 0) > 0 ? 'bg-emerald-100 text-emerald-800 ring-1 ring-emerald-600/20' : 'bg-slate-100 text-slate-400'
                        }`}>{item.quoteCount || 0} Bid(s)</span>
                      </td>
                      <td className="py-4 px-6 text-center">
                        <span className={`inline-flex items-center rounded-full px-2 py-0.5 text-xs font-medium ${
                          item.status === 'Pending' ? 'bg-yellow-50 text-yellow-800 ring-1 ring-yellow-600/20' : 'bg-blue-50 text-blue-800 ring-1 ring-blue-600/20'
                        }`}>{item.status}</span>
                      </td>
                      <td className="py-4 px-6 text-center whitespace-nowrap space-x-1.5">
                        {isEditingRow ? (
                          <>
                            <button onClick={() => handleUpdateItemRow(item.id)} disabled={isSavingCrud} className="rounded bg-emerald-600 px-2.5 py-1 text-xs font-semibold text-white hover:bg-emerald-500">{isSavingCrud ? "Saving..." : "Save"}</button>
                            <button onClick={cancelEditingRow} className="rounded border border-slate-300 bg-white px-2.5 py-1 text-xs font-medium text-slate-700 hover:bg-slate-50">Cancel</button>
                          </>
                        ) : (
                          <>
                            <button onClick={() => openSourcingModal(item)} className="rounded border border-blue-200 bg-blue-50 px-2 py-1 text-xs font-semibold text-blue-700 hover:bg-blue-100">Source</button>
                            <button onClick={() => openQuotesViewerModal(item)} className="rounded border border-slate-200 bg-white px-2 py-1 text-xs font-semibold text-slate-700 hover:bg-slate-50 shadow-sm">Quotes</button>
                            <button onClick={() => startEditingRow(item)} className="rounded border border-slate-300 bg-white px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-100">✏️</button>
                            <button onClick={() => handleDeleteItemRow(item.id)} className="rounded bg-red-50 px-2 py-1 text-xs font-medium text-red-600 hover:bg-red-100">🗑️</button>
                          </>
                        )}
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* DISPATCH SOURCING MODAL */}
      {isModalOpen && selectedItem && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="w-full max-w-lg rounded-lg bg-white p-6 shadow-xl border border-slate-200">
            <div className="border-b border-slate-200 pb-3 mb-4">
              <h3 className="text-lg font-bold text-slate-900">Select Procurement Sources</h3>
              <p className="text-xs text-slate-500 mt-1">Routing Item: <span className="font-mono font-bold text-slate-700">{selectedItem.itemNumber}</span></p>
            </div>
            <div className="max-h-60 overflow-y-auto mb-6 space-y-2 pr-1">
              {suppliers.map((supplier) => (
                <label key={supplier.id} className={`flex items-center justify-between p-3 rounded-md border text-sm cursor-pointer ${selectedSupplierNos.includes(supplier.supplierNo) ? 'border-blue-500 bg-blue-50/50' : 'border-slate-200 hover:bg-slate-50'}`}>
                  <div className="flex items-center gap-3">
                    <input type="checkbox" checked={selectedSupplierNos.includes(supplier.supplierNo)} onChange={() => handleToggleSupplierSelection(supplier.supplierNo)} className="h-4 w-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500" />
                    <div>
                      <p className="font-semibold text-slate-900">{supplier.companyName} <span className="font-mono font-bold text-xs text-blue-600 bg-blue-50 px-1.5 py-0.5 rounded ml-2">{supplier.supplierNo}</span></p>
                      <p className="text-xs text-slate-500">{supplier.email}</p>
                    </div>
                  </div>
                </label>
              ))}
            </div>
            <div className="flex justify-end gap-3 border-t border-slate-200 pt-4">
              <button type="button" onClick={() => Map.valueOf()} className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50" onClick={() => setIsModalOpen(false)}>Cancel</button>
              <button type="button" onClick={handleDispatchRFQ} disabled={selectedSupplierNos.length === 0 || isRouting} className="rounded-md bg-blue-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-blue-500 disabled:bg-blue-300">{isRouting ? "Routing RFQs..." : `Dispatch to ${selectedSupplierNos.length} Vendor(s)`}</button>
            </div>
          </div>
        </div>
      )}

      {/* UPGRADED QUOTES AUDIT LOGGER (SHOWS RELEASED & RECEIVED BIDS) */}
      {isQuotesModalOpen && selectedItem && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="w-full max-w-2xl rounded-lg bg-white p-6 shadow-xl border border-slate-200 flex flex-col max-h-[85vh]">
            <div className="border-b border-slate-200 pb-3 mb-4">
              <h3 className="text-lg font-bold text-slate-900">Procurement Audit Log</h3>
              <p className="text-xs text-slate-500 mt-1">Status check for Item Number Reference: <span className="font-mono font-bold text-blue-600 bg-blue-50 px-2 py-0.5 rounded">{selectedItem.itemNumber}</span></p>
            </div>

            <div className="overflow-y-auto flex-1 my-2 space-y-6 pr-1">
              {isQuotesLoading ? (
                <div className="p-12 text-center text-sm text-slate-500">Querying live logs...</div>
              ) : (
                <>
                  {/* SECTION 1: WHO THE LINE WAS RELEASED TO */}
                  <div>
                    <h4 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">1. Released Vendors (Pending Quotes)</h4>
                    <div className="border border-slate-200 rounded-md bg-slate-50/50 overflow-hidden">
                      {releasedVendors.length === 0 ? (
                        <p className="p-4 text-xs text-slate-400 italic bg-white">No pending responses. All released vendors have submitted quotes.</p>
                      ) : (
                        <table className="w-full text-left border-collapse text-xs">
                          <thead className="bg-slate-100 text-slate-600 font-bold border-b border-slate-200">
                            <tr>
                              <th className="py-2 px-4">Supplier Name</th>
                              <th className="py-2 px-4">Vendor Code</th>
                              <th className="py-2 px-4">Current Status</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-200 bg-white text-slate-700">
                            {releasedVendors.map((vendor) => (
                              <tr key={vendor.id}>
                                <td className="py-2.5 px-4 font-semibold text-slate-800">{getSupplierName(vendor.supplierNo)}</td>
                                <td className="py-2.5 px-4 font-mono text-slate-500">{vendor.supplierNo}</td>
                                <td className="py-2.5 px-4"><span className="inline-flex items-center rounded-md bg-yellow-50 px-2 py-0.5 text-xs font-medium text-yellow-800 ring-1 ring-inset ring-yellow-600/20">Pending Response</span></td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      )}
                    </div>
                  </div>

                  {/* SECTION 2: BIDS RECEIVED DETAILS */}
                  <div>
                    <h4 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">2. Received Procurement Quotes</h4>
                    <div className="border border-slate-200 rounded-md overflow-hidden">
                      {receivedQuotes.length === 0 ? (
                        <p className="p-4 text-xs text-slate-400 italic bg-white">No active bids received from vendors yet.</p>
                      ) : (
                        <table className="w-full text-left border-collapse text-xs">
                          <thead className="bg-slate-100 text-slate-600 font-bold border-b border-slate-200">
                            <tr>
                              <th className="py-2 px-4">Supplier Firm Name</th>
                              <th className="py-2 px-4 text-right">Unit Price ($)</th>
                              <th className="py-2 px-4">Lead Time</th>
                              <th className="py-2 px-4">Vendor Notes Reference</th>
                              <th className="py-2 px-4">Quote Date Stamp</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-200 bg-white text-slate-700">
                            {receivedQuotes.map((quote) => (
                              <tr key={quote.id} className="hover:bg-blue-50/20 transition-all">
                                <td className="py-3 px-4 font-semibold text-slate-900">
                                  <div>{getSupplierName(quote.supplierNo)}</div>
                                  <span className="font-mono text-[10px] text-slate-400 font-bold block mt-0.5">Code: {quote.supplierNo}</span>
                                </td>
                                <td className="py-3 px-4 text-right font-bold text-emerald-700 font-mono text-sm">
                                  ${quote.offeredPrice !== null ? quote.offeredPrice.toFixed(2) : "0.00"}
                                </td>
                                <td className="py-3 px-4 font-medium text-slate-800">{quote.leadTime || "—"}</td>
                                <td className="py-3 px-4 text-slate-500 italic max-w-xs truncate" title={quote.supplierNote}>
                                  {quote.supplierNote || <span className="text-slate-300">None attached</span>}
                                </td>
                                <td className="py-3 px-4 font-medium text-slate-600 whitespace-nowrap">
                                  {formatTimestamp(quote.timestamp)}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      )}
                    </div>
                  </div>
                </>
              )}
            </div>

            <div className="flex justify-end border-t border-slate-200 pt-4 mt-4">
              <button 
                type="button" 
                onClick={() => { setIsQuotesModalOpen(false); setSelectedItem(null); }}
                className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-50 transition-all shadow-sm"
              >
                Close Audit View
              </button>
            </div>
          </div>
        </div>
      )}

      {/* FILTER MODAL SLIDEOUT */}
      {isFilterModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="w-full max-w-sm rounded-lg bg-white p-6 shadow-xl border border-slate-200">
            <div className="border-b border-slate-200 pb-3 mb-4 flex justify-between items-center">
              <h3 className="text-md font-bold text-slate-900">Filter Procurement Items</h3>
              <button type="button" onClick={clearFilterFields} className="text-xs text-blue-600 hover:text-blue-800 font-semibold">Reset All</button>
            </div>
            <div className="space-y-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Keyed RFQ ID / Project Reference</label>
                <input type="text" value={filterRfqId} onChange={(e) => setFilterRfqId(e.target.value)} className="w-full text-sm rounded border border-slate-300 px-3 py-2 uppercase font-mono" placeholder="e.g. REQ-001" />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Item # Identifier</label>
                <input type="text" value={filterItemNumber} onChange={(e) => setFilterItemNumber(e.target.value)} className="w-full text-sm rounded border border-slate-300 px-3 py-2 font-mono" placeholder="e.g. 1001-A" />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Material Description Keyword</label>
                <input type="text" value={filterDescription} onChange={(e) => setFilterDescription(e.target.value)} className="w-full text-sm rounded border border-slate-300 px-3 py-2" placeholder="e.g. Steel Plate" />
              </div>
            </div>
            <div className="flex justify-end gap-2 border-t border-slate-200 pt-4 mt-6">
              <button type="button" onClick={() => setIsFilterModalOpen(false)} className="w-full rounded bg-blue-600 py-2 text-center text-sm font-semibold text-white hover:bg-blue-500">Apply Filters ({filteredMaterials.length} Rows)</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}