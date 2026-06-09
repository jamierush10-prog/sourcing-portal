// src/app/(supplier)/supplier-dashboard/page.tsx
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

  // States
  const [isFilterModalOpen, setIsFilterModalOpen] = useState(false);
  const [filterRfqId, setFilterRfqId] = useState("");
  const [filterItemNumber, setFilterItemNumber] = useState("");
  const [filterDescription, setFilterDescription] = useState("");
  const [filterBuyer, setFilterBuyer] = useState("");
  const [editingId, setEditingId] = useState<string | null>(null);
  const [bidPrice, setBidPrice] = useState<string>("");
  const [leadTime, setLeadTime] = useState<string>("");
  const [vendorNotes, setVendorNotes] = useState<string>("");
  const [isSaving, setIsSaving] = useState(false);
  const [isGeneratingProposal, setIsGeneratingProposal] = useState(false);

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

  // Auth & Data Listeners
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

  // Logout Function - Defined properly inside component scope
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

  const handleExportTableToExcel = async () => {
    setIsGeneratingProposal(true);
    try {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("Quote");
      ws.getCell("A2").value = (profile?.companyName || "VENDOR").toUpperCase();
      ws.getRow(11).values = ["RFQ ID", "Item #", "Description", "Qty", "Price", "Total"];
      
      filteredRows.forEach((item, i) => {
        const row = ws.addRow([item.rfqId, item.itemNumber, item.description, item.quantity, item.offeredPrice || 0, { formula: `=D${12+i}*E${12+i}` }]);
        row.eachCell((c) => c.alignment = { vertical: "middle", horizontal: "center" });
      });

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const a = document.createElement("a");
      a.href = window.URL.createObjectURL(blob);
      a.download = "Quote_Proposal.xlsx";
      a.click();
    } finally { setIsGeneratingProposal(false); }
  };

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <header className="mb-8 flex justify-between items-start border-b pb-5">
        <div>
          <h1 className="text-3xl font-bold">{profile?.companyName || "Vendor"}</h1>
          <button onClick={handleSupplierLogout} className="text-red-600 text-xs font-bold mt-2 hover:underline">🚪 Logout</button>
        </div>
        <div className="text-right text-xs text-slate-500">
          <p className="font-bold text-slate-900">Austal USA</p>
          <p>100 Austal Way, Mobile, AL</p>
        </div>
      </header>

      <div className="flex gap-2 mb-6">
        <button onClick={() => setIsFilterModalOpen(true)} className="bg-white border px-3 py-1 rounded text-sm">🔍 Filter</button>
        <button onClick={handleExportTableToExcel} className="bg-blue-600 text-white px-4 py-1 rounded text-sm font-bold">📑 Generate Quote Proposal</button>
      </div>

      <table className="w-full bg-white border rounded">
        <thead>
          <tr className="bg-slate-100 text-xs text-left">
            <th className="p-3">RFQ ID</th>
            <th className="p-3">Item #</th>
            <th className="p-3">Description</th>
            <th className="p-3">Qty</th>
            <th className="p-3">Price</th>
            <th className="p-3">Actions</th>
          </tr>
        </thead>
        <tbody className="divide-y">
          {filteredRows.map((item) => (
            <tr key={item.id} className={item.isHot ? "bg-red-50" : ""}>
              <td className="p-3">{item.isHot && "🔥 "}{item.rfqId}</td>
              <td className="p-3">{item.itemNumber}</td>
              <td className="p-3">{item.description}</td>
              <td className="p-3">{item.quantity}</td>
              <td className="p-3">${item.offeredPrice?.toFixed(2) || "0.00"}</td>
              <td className="p-3">
                <button onClick={() => setEditingId(item.id)} className="text-blue-600 font-bold text-xs">Edit</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      
      {isFilterModalOpen && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center">
            <div className="bg-white p-6 rounded w-80">
                <input className="w-full border p-2 mb-2" placeholder="RFQ ID..." onChange={(e) => setFilterRfqId(e.target.value)} />
                <button className="w-full bg-blue-600 text-white p-2" onClick={() => setIsFilterModalOpen(false)}>Apply</button>
            </div>
        </div>
      )}
    </div>
  );
}