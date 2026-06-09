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
  const [isDataLoading, setIsDataLoading] = useState(false);

  // Filter States
  const [isFilterModalOpen, setIsFilterModalOpen] = useState(false);
  const [filterRfqId, setFilterRfqId] = useState("");
  const [filterItemNumber, setFilterItemNumber] = useState("");
  const [filterDescription, setFilterDescription] = useState("");

  const [editingId, setEditingId] = useState<string | null>(null);
  const [bidPrice, setBidPrice] = useState("");
  const [isSaving, setIsSaving] = useState(false);

  // Live filter computation
  const filteredRows = useMemo(() => {
    return rfqs.filter((item) => {
      return (
        (item.rfqId || "").toLowerCase().includes(filterRfqId.trim().toLowerCase()) &&
        (item.itemNumber || "").toLowerCase().includes(filterItemNumber.trim().toLowerCase()) &&
        (item.description || "").toLowerCase().includes(filterDescription.trim().toLowerCase())
      );
    });
  }, [rfqs, filterRfqId, filterItemNumber, filterDescription]);

  useEffect(() => {
    if (!loading && (!profile || profile.role !== "supplier")) router.push("/login");
  }, [profile, loading, router]);

  useEffect(() => {
    const supplierProfile = profile as any;
    if (loading || !supplierProfile?.supplierNo) return;
    setIsDataLoading(true);

    const q = query(collection(db, "rfq_routing"), where("supplierNo", "==", supplierProfile.supplierNo));
    const unsubscribe = onSnapshot(q, (snap) => {
      const list: RFQItem[] = [];
      snap.forEach((d) => { list.push({ id: d.id, ...d.data() } as RFQItem); });
      setRfqs(list);
      setIsDataLoading(false);
    });
    return () => unsubscribe();
  }, [profile, loading]);

  const handleExport = async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Quote Proposal");
    ws.getRow(1).values = ["RFQ ID", "Item #", "Description", "Qty", "Price"];
    filteredRows.forEach((item) => {
      ws.addRow([item.rfqId, item.itemNumber, item.description, item.quantity, item.offeredPrice]);
    });
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = document.createElement("a");
    a.href = window.URL.createObjectURL(blob);
    a.download = "Quote_Proposal.xlsx";
    a.click();
  };

  return (
    <div className="min-h-screen bg-slate-50 p-6">
      {/* HEADER */}
      <header className="flex justify-between items-end border-b-2 border-slate-300 pb-6 mb-8">
        <div>
          <h1 className="text-4xl font-black text-slate-900">{profile?.companyName || "Supplier"}</h1>
          <p className="text-slate-600 font-bold mt-1">Bidding Terminal</p>
          <button onClick={() => signOut(auth)} className="text-red-700 text-xs font-bold mt-4 hover:underline">🚪 LOGOUT</button>
        </div>
        <div className="text-right text-xs font-bold text-slate-500">
          <p>AUSTAL USA PROCUREMENT</p>
          <p>100 AUSTAL WAY, MOBILE, AL</p>
        </div>
      </header>

      {/* OPERATIONS BAR */}
      <div className="flex gap-4 mb-6">
        <button onClick={() => setIsFilterModalOpen(true)} className="bg-white border-2 border-slate-400 px-6 py-2 font-bold text-sm shadow-sm">🔍 FILTER DATA</button>
        <button onClick={handleExport} className="bg-blue-900 text-white px-6 py-2 font-bold text-sm shadow-sm">📑 EXPORT TO EXCEL</button>
      </div>

      {/* DATA GRID */}
      <div className="bg-white border-2 border-slate-400 rounded-lg shadow-xl overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-slate-900 text-white text-xs uppercase">
            <tr>
              <th className="p-4">RFQ ID</th>
              <th className="p-4">Item #</th>
              <th className="p-4">Description</th>
              <th className="p-4">Qty</th>
              <th className="p-4">Price</th>
              <th className="p-4">Actions</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-300">
            {filteredRows.map((item) => (
              <tr key={item.id} className={item.isHot ? "bg-red-100 font-bold" : ""}>
                <td className="p-4 font-mono text-slate-900">{item.isHot && "🔥 "}{item.rfqId}</td>
                <td className="p-4 text-slate-900">{item.itemNumber}</td>
                <td className="p-4 text-slate-900">{item.description}</td>
                <td className="p-4 text-slate-900">{item.quantity}</td>
                <td className="p-4 font-bold text-emerald-900">${item.offeredPrice?.toFixed(2) || "0.00"}</td>
                <td className="p-4"><button onClick={() => setEditingId(item.id)} className="text-blue-800 font-bold">EDIT</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* FILTER MODAL */}
      {isFilterModalOpen && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center">
            <div className="bg-white p-8 border-4 border-slate-900 w-96">
                <h3 className="font-black text-xl mb-4">FILTER QUEUE</h3>
                <input className="w-full border-2 p-2 mb-2" placeholder="RFQ ID" onChange={(e) => setFilterRfqId(e.target.value)} />
                <input className="w-full border-2 p-2 mb-2" placeholder="Item #" onChange={(e) => setFilterItemNumber(e.target.value)} />
                <input className="w-full border-2 p-2 mb-4" placeholder="Description" onChange={(e) => setFilterDescription(e.target.value)} />
                <button className="w-full bg-slate-900 text-white py-3 font-bold" onClick={() => setIsFilterModalOpen(false)}>APPLY</button>
            </div>
        </div>
      )}
    </div>
  );
}