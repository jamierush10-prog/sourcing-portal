"use client";

import React, { useState, useEffect, useMemo } from "react";
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

  // Filter States
  const [isFilterModalOpen, setIsFilterModalOpen] = useState(false);
  const [filterRfqId, setFilterRfqId] = useState("");
  const [filterItemNumber, setFilterItemNumber] = useState("");
  const [filterDescription, setFilterDescription] = useState("");
  const [filterBuyer, setFilterBuyer] = useState("");

  // Bidding States
  const [editingId, setEditingId] = useState<string | null>(null);
  const [bidPrice, setBidPrice] = useState<string>("");
  const [leadTime, setLeadTime] = useState<string>("");
  const [vendorNotes, setVendorNotes] = useState<string>("");
  const [isSaving, setIsSaving] = useState(false);
  const [isGeneratingProposal, setIsGeneratingProposal] = useState(false);

  // Live filter computation
  const filteredRows = useMemo(() => {
    return rfqs.filter((item) => {
      return (
        (item.rfqId || "").toLowerCase().includes(filterRfqId.trim().toLowerCase()) &&
        (item.itemNumber || "").toLowerCase().includes(filterItemNumber.trim().toLowerCase()) &&
        (item.description || "").toLowerCase().includes(filterDescription.trim().toLowerCase()) &&
        (item.buyer || "").toLowerCase().includes(filterBuyer.trim().toLowerCase())
      );
    });
  }, [rfqs, filterRfqId, filterItemNumber, filterDescription, filterBuyer]);

  // Auth & Sync
  useEffect(() => {
    if (!loading && (!profile || profile.role !== "supplier")) router.push("/login");
  }, [profile, loading, router]);

  useEffect(() => {
    const supplierProfile = profile as any;
    if (loading || profile?.role !== "supplier" || !supplierProfile?.supplierNo) return;
    setIsDataLoading(true);

    getDocs(collection(db, "materials")).then((snap) => {
      const matMap: Record<string, any> = {};
      snap.forEach((d) => { matMap[d.id] = d.data(); });
      setMaterialsMap(matMap);
    }).catch(console.error);

    const q = query(collection(db, "rfq_routing"), where("supplierNo", "==", supplierProfile.supplierNo));
    const unsubscribe = onSnapshot(q, (snap) => {
      const list: RFQItem[] = [];
      snap.forEach((d) => { list.push({ id: d.id, ...d.data() } as RFQItem); });
      setRfqs(list);
      setIsDataLoading(false);
    });
    return () => unsubscribe();
  }, [profile, loading]);

  const handleSupplierLogout = async () => {
    await signOut(auth);
    router.push("/login");
  };

  const handleSaveBid = async (rfqId: string) => {
    setIsSaving(true);
    try {
      await updateDoc(doc(db, "rfq_routing", rfqId), {
        offeredPrice: parseFloat(bidPrice) || 0,
        leadTime: leadTime.trim(),
        supplierNote: vendorNotes.trim(),
        status: "Completed",
        timestamp: new Date()
      });
      setEditingId(null);
    } finally { setIsSaving(false); }
  };

  const handleExportProposal = async () => {
    setIsGeneratingProposal(true);
    try {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("Quote Proposal");
      
      // Header Section
      ws.mergeCells("A2:C2");
      ws.getCell("A2").value = (profile?.companyName || "VENDOR").toUpperCase();
      ws.getCell("A2").font = { size: 18, bold: true, color: { argb: "1E3A8A" } };
      ws.getCell("A4").value = `Supplier ID: ${(profile as any)?.supplierNo || "N/A"}`;
      ws.getCell("A5").value = `Contact: ${(profile as any)?.contactName || "N/A"}`;
      
      ws.getRow(11).values = ["RFQ ID", "Item #", "Description", "Qty", "Price", "Total"];
      
      filteredRows.forEach((item, i) => {
        const row = ws.addRow([item.rfqId, item.itemNumber, item.description, item.quantity, item.offeredPrice || 0, { formula: `=D${12+i}*E${12+i}` }]);
        row.eachCell((c) => c.alignment = { vertical: "middle", horizontal: "center" });
        row.getCell(5).numFmt = "$#,##0.00";
        row.getCell(6).numFmt = "$#,##0.00";
      });

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const a = document.createElement("a");
      a.href = window.URL.createObjectURL(blob);
      a.download = `Quote_${profile?.companyName || "Proposal"}.xlsx`;
      a.click();
    } finally { setIsGeneratingProposal(false); }
  };

  return (
    <div className="min-h-screen bg-slate-100 p-8">
      {/* HEADER COMMAND CENTER */}
      <header className="bg-white border-b-4 border-blue-900 shadow-lg px-8 py-6 mb-8 rounded-t-lg">
        <div className="flex justify-between items-start">
          <div>
            <h1 className="text-4xl font-black text-slate-900 tracking-tighter">{profile?.companyName || "Vendor Dashboard"}</h1>
            <p className="text-blue-900 font-bold tracking-widest text-xs uppercase mt-1">Bidding Terminal Workspace</p>
            <button onClick={handleSupplierLogout} className="text-red-700 text-xs font-bold mt-4 hover:underline">🚪 LOGOUT SECURELY</button>
          </div>
          <div className="bg-slate-900 text-white p-5 rounded border-b-4 border-blue-600 shadow-xl text-xs">
            <p className="font-bold text-blue-400 uppercase tracking-widest">To: Austal USA</p>
            <p className="mt-2 text-slate-300">100 Austal Way</p>
            <p className="text-slate-300">Mobile, AL 36602</p>
          </div>
        </div>
      </header>

      {/* OPERATIONS BAR */}
      <div className="mx-8 mb-6 flex gap-4">
        <button onClick={() => setIsFilterModalOpen(true)} className="bg-white border-2 border-slate-300 px-6 py-2 rounded-lg font-bold text-sm shadow-sm hover:bg-slate-50">🔍 FILTER QUEUE</button>
        <button onClick={handleExportProposal} className="bg-blue-800 text-white px-6 py-2 rounded-lg font-bold text-sm shadow-sm hover:bg-blue-900">📑 GENERATE QUOTE PROPOSAL</button>
      </div>

      {/* DATA GRID */}
      <div className="mx-8 bg-white border border-slate-300 rounded shadow-sm overflow-hidden">
        <table className="w-full text-left text-sm">
          <thead className="bg-slate-800 text-white text-xs uppercase font-bold">
            <tr>
              <th className="p-4">RFQ ID</th>
              <th className="p-4">Item #</th>
              <th className="p-4">Description</th>
              <th className="p-4">Qty</th>
              <th className="p-4">Your Price</th>
              <th className="p-4 text-center">Actions</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200">
            {filteredRows.map((item) => (
              <tr key={item.id} className={`${item.isHot ? "bg-red-50 border-l-4 border-l-red-500" : "hover:bg-slate-50"}`}>
                <td className="p-4 font-mono font-bold text-slate-800">{item.isHot && "🔥 "}{item.rfqId}</td>
                <td className="p-4">{item.itemNumber}</td>
                <td className="p-4">{item.description}</td>
                <td className="p-4">{item.quantity}</td>
                <td className="p-4 font-bold text-emerald-800">${item.offeredPrice?.toFixed(2) || "0.00"}</td>
                <td className="p-4 text-center">
                    <button onClick={() => setEditingId(item.id)} className="text-blue-700 font-bold text-[11px] underline">EDIT BID</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      
      {/* FILTER MODAL */}
      {isFilterModalOpen && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white p-8 rounded-xl shadow-2xl w-96 border border-slate-200">
            <h3 className="font-black text-lg mb-6 uppercase">Procurement Filters</h3>
            <input className="w-full border-2 p-3 mb-4 rounded font-mono" placeholder="RFQ ID Reference..." onChange={(e) => setFilterRfqId(e.target.value)} />
            <button className="w-full bg-slate-900 text-white py-3 rounded font-bold" onClick={() => setIsFilterModalOpen(false)}>Apply Filters</button>
          </div>
        </div>
      )}
    </div>
  );
}