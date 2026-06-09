"use client";

import React, { useState, useEffect, useMemo } from "react";
import ExcelJS from "exceljs";
import { collection, onSnapshot, query, where, doc, updateDoc, getDocs } from "firebase/firestore";
import { db } from "@/lib/firebase";
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
  const [isDataLoading, setIsDataLoading] = useState(false);
  const [isFilterModalOpen, setIsFilterModalOpen] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [isExporting, setIsExporting] = useState(false);

  // States: Filters
  const [filterRfqId, setFilterRfqId] = useState("");
  const [filterItemNumber, setFilterItemNumber] = useState("");
  const [filterDescription, setFilterDescription] = useState("");
  const [filterBuyer, setFilterBuyer] = useState("");

  // States: Bidding & Editing
  const [editingId, setEditingId] = useState<string | null>(null);
  const [bidPrice, setBidPrice] = useState("");
  const [leadTime, setLeadTime] = useState("");
  const [vendorNotes, setVendorNotes] = useState("");

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

  const handleSaveBid = async (id: string) => {
    setIsSaving(true);
    try {
      await updateDoc(doc(db, "rfq_routing", id), {
        offeredPrice: parseFloat(bidPrice) || 0,
        leadTime,
        supplierNote: vendorNotes,
        status: "Completed",
        timestamp: new Date()
      });
      setEditingId(null);
    } catch (err) {
      console.error(err);
      alert("Error saving your bid.");
    } finally {
      setIsSaving(false);
    }
  };

  const handleExport = async () => {
    setIsExporting(true);
    try {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("Quote Proposal");
      
      // Original detailed column structure
      ws.columns = [
        { header: "RFQ ID", key: "rfqId", width: 15 },
        { header: "Item #", key: "item", width: 12 },
        { header: "Description", key: "desc", width: 40 },
        { header: "Quantity", key: "qty", width: 10 },
        { header: "UOM", key: "uom", width: 8 },
        { header: "Buyer", key: "buyer", width: 20 },
        { header: "Your Price ($)", key: "price", width: 18 },
        { header: "Lead Time", key: "leadTime", width: 16 },
        { header: "Notes", key: "notes", width: 30 }
      ];

      filteredRows.forEach((item) => {
        const row = ws.addRow({
            rfqId: item.rfqId,
            item: item.itemNumber,
            desc: item.description,
            qty: item.quantity,
            uom: item.uom,
            buyer: item.buyer,
            price: item.offeredPrice,
            leadTime: item.leadTime,
            notes: item.supplierNote
        });
        row.eachCell((c) => c.alignment = { vertical: "middle", horizontal: "center" });
      });

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const a = document.createElement("a");
      a.href = window.URL.createObjectURL(blob);
      a.download = "Quote_Proposal.xlsx";
      a.click();
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="min-h-screen bg-white p-8 text-black">
      {/* HEADER SECTION - MAINTAINED INDUSTRIAL LAYOUT */}
      <header className="border-b-4 border-black pb-8 mb-8 flex justify-between items-end">
        <div>
          <h1 className="text-5xl font-black uppercase tracking-tighter text-black">{profile?.companyName || "Vendor Dashboard"}</h1>
          <p className="text-xs font-black text-black uppercase mt-2 tracking-widest">Bidding Terminal | Shipyard Procurement Workspace</p>
          <button onClick={() => signOut(auth)} className="mt-6 border-2 border-black px-6 py-2 font-black text-xs uppercase hover:bg-black hover:text-white transition-all">Secure Logout</button>
        </div>
        <div className="text-right border-l-4 border-black pl-6">
          <p className="font-black text-2xl uppercase">Austal USA</p>
          <p className="font-bold text-xs uppercase">100 Austal Way, Mobile, AL</p>
        </div>
      </header>

      {/* OPERATIONS TOOLBAR */}
      <div className="flex gap-4 mb-8">
        <button onClick={() => setIsFilterModalOpen(true)} className="bg-black text-white px-8 py-3 font-black text-xs uppercase hover:bg-slate-800">Filter Data</button>
        <button onClick={handleExport} className="border-2 border-black px-8 py-3 font-black text-xs uppercase hover:bg-black hover:text-white">Export to Excel</button>
      </div>

      {/* DATA GRID - HIGH VISIBILITY, FULL DENSITY */}
      <div className="border-2 border-black shadow-[8px_8px_0px_0px_rgba(0,0,0,1)]">
        <table className="w-full text-left text-sm text-black">
          <thead className="bg-black text-white uppercase font-black text-xs border-b-2 border-black">
            <tr>
              <th className="p-4 border-r border-slate-700">RFQ ID</th>
              <th className="p-4 border-r border-slate-700">Item #</th>
              <th className="p-4 border-r border-slate-700">Description</th>
              <th className="p-4 border-r border-slate-700">Qty</th>
              <th className="p-4 border-r border-slate-700">Buyer</th>
              <th className="p-4 border-r border-slate-700">Price</th>
              <th className="p-4">Actions</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-black">
            {filteredRows.map((item) => (
              <tr key={item.id} className="hover:bg-slate-100 transition-colors">
                <td className="p-4 border-r border-black font-mono font-black">{item.rfqId}</td>
                <td className="p-4 border-r border-black font-bold">{item.itemNumber}</td>
                <td className="p-4 border-r border-black font-bold">{item.description}</td>
                <td className="p-4 border-r border-black font-bold">{item.quantity}</td>
                <td className="p-4 border-r border-black font-bold">{item.buyer}</td>
                <td className="p-4 border-r border-black font-black text-emerald-800">${item.offeredPrice?.toFixed(2) || "0.00"}</td>
                <td className="p-4 text-center">
                    <button onClick={() => setEditingId(item.id)} className="font-black text-xs underline">EDIT BID</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* FILTER MODAL - INDUSTRIAL STYLE */}
      {isFilterModalOpen && (
        <div className="fixed inset-0 bg-black/80 flex items-center justify-center p-6">
            <div className="bg-white p-10 border-4 border-black w-full max-w-xl shadow-[16px_16px_0px_0px_rgba(0,0,0,1)]">
                <h3 className="font-black text-2xl uppercase mb-8">Filter Data Parameters</h3>
                <div className="grid grid-cols-2 gap-6">
                    <input className="border-2 border-black p-4 font-bold" placeholder="RFQ ID..." onChange={(e) => setFilterRfqId(e.target.value)} />
                    <input className="border-2 border-black p-4 font-bold" placeholder="Item #..." onChange={(e) => setFilterItemNumber(e.target.value)} />
                    <input className="border-2 border-black p-4 font-bold col-span-2" placeholder="Description Keyword..." onChange={(e) => setFilterDescription(e.target.value)} />
                    <input className="border-2 border-black p-4 font-bold col-span-2" placeholder="Buyer Name..." onChange={(e) => setFilterBuyer(e.target.value)} />
                </div>
                <button className="w-full bg-black text-white py-5 font-black uppercase mt-10 hover:bg-slate-800" onClick={() => setIsFilterModalOpen(false)}>Apply Active Filters</button>
            </div>
        </div>
      )}
    </div>
  );
}