// src/app/(admin)/dashboard/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import Papa from "papaparse";
import ExcelJS from "exceljs";
import { collection, writeBatch, doc, onSnapshot, query, orderBy, where, updateDoc, getDocs } from "firebase/firestore";
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
  buyer: string; 
  status: "Pending" | "Sourced" | "Completed";
  isHot?: boolean;
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
  buyer: string;
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

  // Bulk Row Checkbox Selection State
  const [selectedItemIds, setSelectedItemIds] = useState<string[]>([]);

  // Filter Modal Parameter States
  const [isFilterModalOpen, setIsFilterModalOpen] = useState(false);
  const [filterRfqId, setFilterRfqId] = useState("");
  const [filterItemNumber, setFilterItemNumber] = useState("");
  const [filterDescription, setFilterDescription] = useState("");
  const [filterBuyer, setFilterBuyer] = useState("");

  // CRUD Inline Editing States
  const [editingItemId, setEditingItemId] = useState<string | null>(null);
  const [editRfqId, setEditRfqId] = useState("");
  const [editItemNumber, setEditItemNumber] = useState("");
  const [editDescription, setEditDescription] = useState("");
  const [editQuantity, setEditQuantity] = useState<number>(0);
  const [editUom, setEditUom] = useState("");
  const [editBuyer, setEditBuyer] = useState("");
  const [isSavingCrud, setIsSavingCrud] = useState(false);

  // Sourcing Modal States (Handles both single and bulk actions)
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [selectedSupplierNos, setSelectedSupplierNos] = useState<string[]>([]);
  const [isRouting, setIsRouting] = useState(false);

  // Quotes Viewer Modal States
  const [isQuotesModalOpen, setIsQuotesModalOpen] = useState(false);
  const [selectedItem, setSelectedItem] = useState<MaterialItem | null>(null);
  const [releasedVendors, setReleasedVendors] = useState<BidResponse[]>([]);
  const [receivedQuotes, setReceivedQuotes] = useState<BidResponse[]>([]);
  const [isQuotesLoading, setIsQuotesLoading] = useState(false);

  const [isExportingExcel, setIsExportingExcel] = useState(false);

  useEffect(() => {
    if (!loading && (!profile || profile.role !== "admin")) {
      router.push("/login");
    }
  }, [profile, loading, router]);

  useEffect(() => {
    if (profile?.role !== "admin") return;

    fetchSuppliers();
    const mQuery = query(collection(db, "materials"), orderBy("itemNumber", "asc"));
    
    const unsubscribeMaterials = onSnapshot(mQuery, async (materialsSnapshot) => {
      const itemsList: MaterialItem[] = [];
      materialsSnapshot.forEach((doc) => {
        itemsList.push({ id: doc.id, ...doc.data(), quoteCount: 0 } as MaterialItem);
      });

      try {
        const routingQuery = query(collection(db, "rfq_routing"), where("status", "==", "Completed"));
        const routingSnapshot = await getDocs(routingQuery);
        
        routingSnapshot.forEach((routingDoc) => {
          const data = routingDoc.data();
          const matchingMaterial = itemsList.find(item => item.id === data.materialId);
          if (matchingMaterial && matchingMaterial.quoteCount !== undefined) {
            matchingMaterial.quoteCount += 1;
          }
        });
      } catch (err) {
        console.error("Error cross-referencing bid counts:", err);
      }

      setMaterials(itemsList);
    }, (err) => {
      console.error("Live materials stream failed:", err);
    });

    return () => unsubscribeMaterials();
  }, [profile]);

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

  // Checkbox Row Selection Logic
  const handleToggleRowSelect = (id: string) => {
    setSelectedItemIds(prev => 
      prev.includes(id) ? prev.filter(item => item !== id) : [...prev, id]
    );
  };

  const handleSelectAllVisible = (visibleItems: MaterialItem[]) => {
    const visibleIds = visibleItems.map(item => item.id);
    const allSelected = visibleIds.every(id => selectedItemIds.includes(id));
    
    if (allSelected) {
      setSelectedItemIds(prev => prev.filter(id => !visibleIds.includes(id)));
    } else {
      setSelectedItemIds(prev => Array.from(new Set([...prev, ...visibleIds])));
    }
  };

  // RECONCILED: Modal Supplier Selection Toggle Logic Method
  const handleToggleSupplierSelection = (supplierNo: string) => {
    setSelectedSupplierNos((prev) =>
      prev.includes(supplierNo) ? prev.filter((no) => no !== supplierNo) : [...prev, supplierNo]
    );
  };

  // BULK ACTION: FLAG SELECTED ITEMS AS HOT
  const handleMarkSelectedHot = async () => {
    if (selectedItemIds.length === 0) {
      alert("Please select at least one material item line.");
      return;
    }

    try {
      const batch = writeBatch(db);
      selectedItemIds.forEach((id) => {
        const docRef = doc(db, "materials", id);
        batch.update(docRef, { isHot: true });

        const rQuery = query(collection(db, "rfq_routing"), where("materialId", "==", id));
        getDocs(rQuery).then((rSnapshot) => {
          const innerBatch = writeBatch(db);
          rSnapshot.forEach((rDoc) => {
            innerBatch.update(rDoc.ref, { isHot: true });
          });
          innerBatch.commit();
        });
      });

      await batch.commit();
      setSelectedItemIds([]);
      setFeedbackMessage(`Flagged ${selectedItemIds.length} item(s) as Critical Hot 🔥`);
      setTimeout(() => setFeedbackMessage(""), 4000);
    } catch (err) {
      console.error("Failed to execute hot state assignment update batch:", err);
    }
  };

  // Open Bulk Source Modal
  const openBulkSourcingModal = () => {
    if (selectedItemIds.length === 0) {
      alert("Please select at least one check box item line.");
      return;
    }
    setSelectedSupplierNos([]);
    setIsModalOpen(true);
  };

  // BULK DISPATCH RFQ ROUTING CONSOLE EXECUTOR
  const handleDispatchBulkRFQs = async () => {
    if (selectedItemIds.length === 0 || selectedSupplierNos.length === 0) return;
    setIsRouting(true);

    try {
      const batch = writeBatch(db);

      selectedItemIds.forEach((itemId) => {
        const matchItem = materials.find(m => m.id === itemId);
        if (!matchItem) return;

        selectedSupplierNos.forEach((supNo) => {
          const rfqDocRef = doc(collection(db, "rfq_routing"));
          batch.set(rfqDocRef, {
            materialId: matchItem.id,
            rfqId: matchItem.rfqId || "—", 
            itemNumber: matchItem.itemNumber,
            description: matchItem.description,
            quantity: matchItem.quantity,
            uom: matchItem.uom,
            buyer: matchItem.buyer || "UNASSIGNED", 
            supplierNo: supNo,
            status: "Pending",
            isHot: matchItem.isHot || false,
            offeredPrice: null,
            leadTime: null,
            supplierNote: "",
            timestamp: null, 
          });
        });

        const materialDocRef = doc(db, "materials", itemId);
        batch.update(materialDocRef, { status: "Sourced" });
      });

      await batch.commit();
      setIsModalOpen(false);
      setSelectedItemIds([]);
      setFeedbackMessage("Bulk routing matrix routed successfully.");
      setTimeout(() => setFeedbackMessage(""), 4000);
    } catch (err) {
      console.error("Bulk RFQ deployment batch assignment failure:", err);
    } finally {
      setIsRouting(false);
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
              buyer: row["Buyer"] || row["buyer"] || "UNASSIGNED", 
              status: "Pending",
              isHot: false,
              timestamp: new Date()
            });
          });
          await batch.commit();
          setFeedbackMessage("Requirements queue initialized.");
        } catch (error) {
          console.error("CSV upload batch failed:", error);
        } finally {
          setIsUploading(false);
        }
      },
    });
  };

  const downloadCsvTemplate = () => {
    const headers = ["RFQ ID", "Item", "Description", "Qty", "UOM", "Buyer"];
    const sampleRows = [
      ["REQ-2026-01", "1001-A", "3/4\" Structural Carbon Steel Plate A36", "12", "EA", "James Rush"],
      ["PROJECT-BLUE", "1002-B", "Grade 60 #5 Rebar Deformed Steel 20ft", "250", "LF", "J. Rush"]
    ];
    const csvContent = [headers, ...sampleRows].map(row => row.map(val => `"${val.replace(/"/g, '""')}"`).join(",")).join("\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.setAttribute("download", "Material_Requirements_Template.csv");
    link.click();
    URL.revokeObjectURL(url);
  };

  const startEditingRow = (item: MaterialItem) => {
    setEditingItemId(item.id);
    setEditRfqId(item.rfqId || "");
    setEditItemNumber(item.itemNumber);
    setEditDescription(item.description);
    setEditQuantity(item.quantity);
    setEditUom(item.uom);
    setEditBuyer(item.buyer || "");
  };

  const handleUpdateItemRow = async (id: string) => {
    setIsSavingCrud(true);
    try {
      const batch = writeBatch(db);
      const docRef = doc(db, "materials", id);
      batch.update(docRef, {
        rfqId: editRfqId.trim(), 
        itemNumber: editItemNumber.trim(),
        description: editDescription.trim(),
        quantity: Number(editQuantity),
        uom: editUom.trim(),
        buyer: editBuyer.trim() 
      });

      const subQuery = query(collection(db, "rfq_routing"), where("materialId", "==", id));
      const subSnapshot = await getDocs(subQuery);
      subSnapshot.forEach((subDoc) => {
        batch.update(subDoc.ref, {
          rfqId: editRfqId.trim(),
          itemNumber: editItemNumber.trim(),
          description: editDescription.trim(),
          quantity: Number(editQuantity),
          uom: editUom.trim(),
          buyer: editBuyer.trim()
        });
      });

      await batch.commit();
      setEditingItemId(null);
    } catch (err) {
      console.error(err);
    } finally {
      setIsSavingCrud(false);
    }
  };

  const handleDeleteItemRow = async (id: string) => {
    if (!confirm("Delete requirement row?")) return;
    try {
      const batch = writeBatch(db);
      batch.delete(doc(db, "materials", id));
      const rSub = query(collection(db, "rfq_routing"), where("materialId", "==", id));
      const rSnapshot = await getDocs(rSub);
      rSnapshot.forEach((rDoc) => batch.delete(rDoc.ref));
      await batch.commit();
    } catch (err) {
      console.error(err);
    }
  };

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
          completedBids.push({ ...data, id: doc.id });
        } else {
          pendingRelease.push({ ...data, id: doc.id });
        }
      });
      completedBids.sort((a, b) => (a.offeredPrice || 0) - (b.offeredPrice || 0));
      setReleasedVendors(pendingRelease);
      setReceivedQuotes(completedBids);
    } catch (err) {
      console.error(err);
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
      snapshot.forEach(doc => allQuotes.push(doc.data()));

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Master Received Quotes Log");
      worksheet.columns = [
        { header: "RFQ ID", key: "rfqId", width: 16 },
        { header: "Item #", key: "itemNo", width: 14 },
        { header: "Material Description", key: "desc", width: 36 },
        { header: "Quantity", key: "qty", width: 10 },
        { header: "UOM", key: "uom", width: 8 },
        { header: "Buyer", key: "buyer", width: 16 }, 
        { header: "Supplier", key: "supplierName", width: 26 },
        { header: "Price ($)", key: "price", width: 20 },
        { header: "Lead Time", key: "leadTime", width: 16 }
      ];

      allQuotes.forEach((quote) => {
        const v = suppliers.find(s => s.supplierNo === quote.supplierNo);
        worksheet.addRow({
          rfqId: quote.rfqId || "—", itemNo: quote.itemNumber || "—", desc: quote.description || "",
          qty: Number(quote.quantity || 0), uom: quote.uom || "EA", buyer: quote.buyer || "—",
          supplierName: v ? v.companyName : quote.supplierNo, price: Number(quote.offeredPrice || 0),
          leadTime: quote.leadTime || "—"
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `Procurement_Master_Quotes_${new Date().toISOString().substring(0,10)}.xlsx`;
      anchor.click();
    } catch (err) {
      console.error(err);
    } finally {
      setIsExportingExcel(false);
    }
  };

  const filteredMaterials = materials.filter((item) => {
    return (
      (item.rfqId || "").toLowerCase().includes(filterRfqId.trim().toLowerCase()) &&
      (item.itemNumber || "").toLowerCase().includes(filterItemNumber.trim().toLowerCase()) &&
      (item.description || "").toLowerCase().includes(filterDescription.trim().toLowerCase()) &&
      (item.buyer || "").toLowerCase().includes(filterBuyer.trim().toLowerCase())
    );
  });

  const clearFilterFields = () => {
    setFilterRfqId("");
    setFilterItemNumber("");
    setFilterDescription("");
    setFilterBuyer("");
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
          <p className="text-sm text-slate-500">Bulk action workspace module logs tracking indexes</p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={openBulkSourcingModal}
            disabled={selectedItemIds.length === 0}
            className="text-sm font-bold text-blue-700 bg-blue-50 border border-blue-200 px-3 py-1.5 rounded-md hover:bg-blue-100 disabled:opacity-40 shadow-sm transition-all"
          >
            🏢 Bulk Source Selected ({selectedItemIds.length})
          </button>
          
          <button
            type="button"
            onClick={handleMarkSelectedHot}
            disabled={selectedItemIds.length === 0}
            className="text-sm font-bold text-red-700 bg-red-50 border border-red-200 px-3 py-1.5 rounded-md hover:bg-red-100 disabled:opacity-40 shadow-sm transition-all"
          >
            🔥 Mark Selected Hot ({selectedItemIds.length})
          </button>

          <button onClick={() => setIsFilterModalOpen(true)} className="text-sm font-semibold text-slate-700 bg-white border border-slate-300 px-3 py-1.5 rounded-md hover:bg-slate-50 shadow-sm transition-all">🔍 Filter Queue</button>
          <button onClick={handleExportAllQuotesToExcel} disabled={isExportingExcel} className="text-sm font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-3 py-1.5 rounded-md hover:bg-emerald-100 shadow-sm">{isExportingExcel ? "Exporting..." : "📊 Export All Quotes"}</button>
          <button onClick={() => router.push("/suppliers")} className="text-sm font-semibold text-blue-600 hover:text-blue-800 bg-blue-50 border border-blue-200 px-3 py-1.5 rounded-md transition-all">🏢 Suppliers</button>
        </div>
      </header>

      {/* CSV Import Layout */}
      <div className="mb-8 p-4 bg-white border border-slate-200 rounded-lg shadow-sm flex flex-col sm:flex-row gap-4 justify-between items-start sm:items-center">
        <div>
          <h4 className="text-sm font-bold text-slate-800">Import Master Material Requirements</h4>
          <p className="text-xs text-slate-500">Upload bulk CSV lists with automated parser mappings</p>
        </div>
        <div className="flex flex-wrap items-center gap-3 w-full sm:w-auto justify-end">
          <button type="button" onClick={downloadCsvTemplate} className="text-xs font-bold text-blue-700 bg-blue-50 border border-blue-200 px-3 py-1.5 rounded transition-colors">📥 Template</button>
          <input type="file" accept=".csv" onChange={handleCSVUpload} disabled={isUploading} className="text-xs text-slate-500 file:py-1 file:px-3 file:rounded file:border-0 file:bg-blue-50 file:text-blue-700 cursor-pointer" />
          {feedbackMessage && <span className="text-xs font-bold text-blue-700 bg-blue-50 border border-blue-200 px-2.5 py-1 rounded shadow-sm">{feedbackMessage}</span>}
        </div>
      </div>

      {/* Requirements Table Card Grid Stream Container */}
      <div className="bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-sm">
            <thead className="bg-slate-100 text-slate-700 font-semibold text-xs border-b border-slate-200">
              <tr>
                <th className="py-3 px-4 text-center w-12">
                  <input 
                    type="checkbox" 
                    checked={filteredMaterials.length > 0 && filteredMaterials.every(m => selectedItemIds.includes(m.id))}
                    onChange={() => handleSelectAllVisible(filteredMaterials)}
                    className="h-4 w-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500 cursor-pointer"
                  />
                </th>
                <th className="py-3 px-4 text-center">RFQ ID</th>
                <th className="py-3 px-6">Item #</th>
                <th className="py-3 px-6">Description</th>
                <th className="py-3 px-6 text-right">Qty</th>
                <th className="py-3 px-6">UOM</th>
                <th className="py-3 px-6">Buyer</th> 
                <th className="py-3 px-6 text-center">Bids</th>
                <th className="py-3 px-6 text-center">Status</th>
                <th className="py-3 px-6 text-center">Operations</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 text-slate-800">
              {filteredMaterials.length === 0 ? (
                <tr><td colSpan={10} className="py-8 text-center text-slate-400">No matching requirements located.</td></tr>
              ) : (
                filteredMaterials.map((item) => {
                  const isEditingRow = editingItemId === item.id;
                  const isChecked = selectedItemIds.includes(item.id);

                  return (
                    <tr key={item.id} className={`hover:bg-slate-50/50 transition-colors ${isEditingRow ? 'bg-amber-50/40' : ''} ${isChecked ? 'bg-blue-50/20' : ''}`}>
                      <td className="py-4 px-4 text-center">
                        <input 
                          type="checkbox" 
                          checked={isChecked}
                          onChange={() => handleToggleRowSelect(item.id)}
                          className="h-4 w-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500 cursor-pointer"
                        />
                      </td>
                      <td className="py-4 px-4 text-center font-mono font-bold text-xs text-slate-700 bg-slate-50/20">
                        {isEditingRow ? (
                          <input type="text" value={editRfqId} onChange={(e) => setEditRfqId(e.target.value)} className="w-24 text-center text-xs rounded border border-slate-300 px-1 py-1 text-slate-900" />
                        ) : (
                          <span className="flex items-center justify-center gap-1">
                            {item.isHot && <span title="Priority Item" className="text-red-500">🔥</span>}
                            {item.rfqId || "—"}
                          </span>
                        )}
                      </td>
                      <td className="py-4 px-6 font-mono font-medium text-slate-900">{isEditingRow ? <input type="text" value={editItemNumber} onChange={(e) => setEditItemNumber(e.target.value)} className="w-28 text-sm rounded border border-slate-300 px-2 py-1 text-slate-900" /> : item.itemNumber}</td>
                      <td className="py-4 px-6 max-w-xs truncate" title={item.description}>{isEditingRow ? <input type="text" value={editDescription} onChange={(e) => setEditDescription(e.target.value)} className="w-full min-w-[180px] text-sm rounded border border-slate-300 px-2 py-1 text-slate-900" /> : item.description}</td>
                      <td className="py-4 px-6 text-right font-semibold">{isEditingRow ? <input type="number" value={editQuantity} onChange={(e) => setEditQuantity(Number(e.target.value))} className="w-16 text-sm text-right rounded border border-slate-300 px-2 py-1 text-slate-900" /> : item.quantity}</td>
                      <td className="py-4 px-6 text-slate-500">{isEditingRow ? <input type="text" value={editUom} onChange={(e) => setEditUom(e.target.value)} className="w-14 text-sm rounded border border-slate-300 px-2 py-1 text-slate-900" /> : item.uom}</td>
                      <td className="py-4 px-6 text-slate-700 font-medium">{isEditingRow ? <input type="text" value={editBuyer} onChange={(e) => setEditBuyer(e.target.value)} className="w-24 text-sm rounded border border-slate-300 px-2 py-1 text-slate-900" /> : item.buyer || "—"}</td>
                      <td className="py-4 px-6 text-center"><span className={`inline-flex items-center rounded-md px-2.5 py-0.5 text-xs font-bold ${(item.quoteCount || 0) > 0 ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-100 text-slate-400'}`}>{item.quoteCount || 0} Bids</span></td>
                      <td className="py-4 px-6 text-center"><span className={`inline-flex items-center rounded-full px-2 py-0.5 text-xs font-medium ${item.status === 'Pending' ? 'bg-yellow-50 text-yellow-800' : 'bg-blue-50 text-blue-800'}`}>{item.status}</span></td>
                      <td className="py-4 px-6 text-center whitespace-nowrap space-x-1.5">
                        {isEditingRow ? (
                          <>
                            <button onClick={() => handleUpdateItemRow(item.id)} disabled={isSavingCrud} className="rounded bg-emerald-600 px-2.5 py-1 text-xs font-bold text-white hover:bg-emerald-500">{isSavingCrud ? "Saving..." : "Save"}</button>
                            <button onClick={() => setEditingItemId(null)} className="rounded border border-slate-300 bg-white px-2.5 py-1 text-xs text-slate-700 hover:bg-slate-50">Cancel</button>
                          </>
                        ) : (
                          <>
                            <button onClick={() => { setSelectedItemIds([item.id]); openBulkSourcingModal(); }} className="rounded border border-blue-200 bg-blue-50 px-2 py-1 text-xs font-semibold text-blue-700 hover:bg-blue-100">Source</button>
                            <button onClick={() => openQuotesViewerModal(item)} className="rounded border border-slate-200 bg-white px-2 py-1 text-xs font-semibold text-slate-700 hover:bg-slate-50">Quotes</button>
                            <button onClick={() => startEditingRow(item)} className="rounded border border-slate-300 bg-white px-1.5 py-1 text-xs text-slate-600 hover:bg-slate-100">✏️</button>
                            <button onClick={() => handleDeleteItemRow(item.id)} className="rounded bg-red-50 px-1.5 py-1 text-xs text-red-600 hover:bg-red-100">🗑️</button>
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

      {/* DISPATCH SOURCING MODAL (SUPPORTS SINGLE & BULK ACTIONS) */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="w-full max-w-lg rounded-lg bg-white p-6 shadow-xl border border-slate-200">
            <div className="border-b border-slate-200 pb-3 mb-4">
              <h3 className="text-lg font-bold text-slate-900">Select Procurement Sources</h3>
              <p className="text-xs text-slate-500 mt-1">Routing Matrix Strategy Allocation: <span className="font-bold text-blue-600">Selected Queue Items Stack Block</span></p>
            </div>
            <div className="max-h-60 overflow-y-auto mb-6 space-y-2 pr-1">
              {suppliers.map((supplier) => (
                <label key={supplier.id} className={`flex items-center justify-between p-3 rounded-md border text-sm cursor-pointer ${selectedSupplierNos.includes(supplier.supplierNo) ? 'border-blue-500 bg-blue-50/50' : 'border-slate-200 hover:bg-slate-50'}`}>
                  <div className="flex items-center gap-3">
                    <input 
                      type="checkbox" 
                      checked={selectedSupplierNos.includes(supplier.supplierNo)} 
                      onChange={() => handleToggleSupplierSelection(supplier.supplierNo)} 
                      className="h-4 w-4 rounded border-slate-300 text-blue-600 cursor-pointer" 
                    />
                    <div>
                      <p className="font-semibold text-slate-900">{supplier.companyName} <span className="font-mono font-bold text-xs text-blue-600 bg-blue-50 px-1.5 py-0.5 rounded ml-2">{supplier.supplierNo}</span></p>
                      <p className="text-xs text-slate-500">{supplier.email}</p>
                    </div>
                  </div>
                </label>
              ))}
            </div>
            <div className="flex justify-end gap-3 border-t border-slate-200 pt-4">
              <button type="button" className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50" onClick={() => setIsModalOpen(false)}>Cancel</button>
              <button type="button" onClick={handleDispatchBulkRFQs} disabled={selectedSupplierNos.length === 0 || isRouting} className="rounded-md bg-blue-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-blue-500 disabled:bg-blue-300">{isRouting ? "Routing Matrix..." : `Dispatch RFQs`}</button>
            </div>
          </div>
        </div>
      )}

      {/* RECEIVED QUOTES AUDIT LOGGER */}
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
                  <div>
                    <h4 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">1. Released Vendors (Pending Quotes)</h4>
                    <div className="border border-slate-200 rounded-md bg-slate-50/50 overflow-hidden">
                      {releasedVendors.length === 0 ? (
                        <p className="p-4 text-xs text-slate-400 italic bg-white">No pending responses.</p>
                      ) : (
                        <table className="w-full text-left border-collapse text-xs">
                          <tbody className="divide-y divide-slate-200 bg-white text-slate-700">
                            {releasedVendors.map((vendor) => (
                              <tr key={vendor.id}>
                                <td className="py-2.5 px-4 font-semibold text-slate-800">{getSupplierName(vendor.supplierNo)}</td>
                                <td className="py-2.5 px-4 font-mono text-slate-500">{vendor.supplierNo}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      )}
                    </div>
                  </div>
                  <div>
                    <h4 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">2. Received Procurement Quotes</h4>
                    <div className="border border-slate-200 rounded-md overflow-hidden">
                      {receivedQuotes.length === 0 ? (
                        <p className="p-4 text-xs text-slate-400 italic bg-white">No active bids received yet.</p>
                      ) : (
                        <table className="w-full text-left border-collapse text-xs">
                          <tbody className="divide-y divide-slate-200 bg-white text-slate-700">
                            {receivedQuotes.map((quote) => (
                              <tr key={quote.id}>
                                <td className="py-3 px-4 font-semibold text-slate-900">{getSupplierName(quote.supplierNo)}</td>
                                <td className="py-3 px-4 text-right font-bold text-emerald-700">${quote.offeredPrice !== null ? quote.offeredPrice.toFixed(2) : "0.00"}</td>
                                <td className="py-3 px-4 font-medium">{quote.leadTime || "—"}</td>
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
              <button type="button" onClick={() => setIsQuotesModalOpen(false)} className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-50">Close</button>
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
              <div><label className="block text-xs font-semibold text-slate-700 mb-1">RFQ ID</label><input type="text" value={filterRfqId} onChange={(e) => setFilterRfqId(e.target.value)} className="w-full text-sm rounded border border-slate-300 px-3 py-2 text-slate-900" placeholder="e.g. REQ-001" /></div>
              <div><label className="block text-xs font-semibold text-slate-700 mb-1">Item #</label><input type="text" value={filterItemNumber} onChange={(e) => setFilterItemNumber(e.target.value)} className="w-full text-sm rounded border border-slate-300 px-3 py-2 text-slate-900" placeholder="e.g. 1001-A" /></div>
              <div><label className="block text-xs font-semibold text-slate-700 mb-1">Description Keyword</label><input type="text" value={filterDescription} onChange={(e) => setFilterDescription(e.target.value)} className="w-full text-sm rounded border border-slate-300 px-3 py-2 text-slate-900" placeholder="e.g. Steel Plate" /></div>
              <div><label className="block text-xs font-semibold text-slate-700 mb-1">Buyer Assigned</label><input type="text" value={filterBuyer} onChange={(e) => setFilterBuyer(e.target.value)} className="w-full text-sm rounded border border-slate-300 px-3 py-2 text-slate-900" placeholder="e.g. James Rush" /></div>
            </div>
            <div className="flex justify-end gap-2 border-t border-slate-200 pt-4 mt-6">
              <button type="button" onClick={() => setIsFilterModalOpen(false)} className="w-full rounded bg-blue-600 py-2 text-center text-sm font-semibold text-white hover:bg-blue-500">Apply Filters</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}