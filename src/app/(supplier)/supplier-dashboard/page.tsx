// src/app/(supplier)/dashboard/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import { collection, getDocs, query, where, doc, updateDoc } from "firebase/firestore";
import { db } from "@/lib/firebase";
import { useAuth } from "@/context/AuthContext";
import { useRouter } from "next/navigation";

interface RFQItem {
  id: string;
  itemNumber: string;
  description: string;
  quantity: number;
  uom: string;
  status: "Pending" | "Responded";
  offeredPrice: number | null;
  leadTime: string | null;
  supplierNote: string;
}

export default function SupplierDashboard() {
  const { profile, loading } = useAuth();
  const router = useRouter();

  const [rfqs, setRfqs] = useState<RFQItem[]>([]);
  const [editingId, setEditingId] = useState<string | null>(null);
  
  // Local form states for bidding editing inline
  const [price, setPrice] = useState<string>("");
  const [leadTime, setLeadTime] = useState<string>("");
  const [note, setNote] = useState<string>("");
  const [isUpdating, setIsUpdating] = useState(false);

  // Authorization routing sentinel
  useEffect(() => {
    if (!loading && (!profile || profile.role !== "supplier")) {
      router.push("/login");
    }
  }, [profile, loading, router]);

  // Read dispatched RFQs owned by current logged-in user profile
  useEffect(() => {
    if (profile?.role === "supplier") {
      fetchSupplierRFQs();
    }
  }, [profile]);

  const fetchSupplierRFQs = async () => {
    if (!profile) return;
    try {
      const q = query(
        collection(db, "rfq_routing"),
        where("supplierId", "==", profile.uid)
      );
      const snapshot = await getDocs(q);
      const list: RFQItem[] = [];
      snapshot.forEach((doc) => {
        list.push({ id: doc.id, ...doc.data() } as RFQItem);
      });
      setRfqs(list);
    } catch (err) {
      console.error("Error pulling isolated vendor payload:", err);
    }
  };

  const startEditing = (item: RFQItem) => {
    setEditingId(item.id);
    setPrice(item.offeredPrice?.toString() || "");
    setLeadTime(item.leadTime || "");
    setNote(item.supplierNote || "");
  };

  const handleUpdateBid = async (id: string) => {
    setIsUpdating(true);
    try {
      const docRef = doc(db, "rfq_routing", id);
      await updateDoc(docRef, {
        offeredPrice: price ? Number(price) : null,
        leadTime: leadTime || null,
        supplierNote: note,
        status: "Responded",
      });
      setEditingId(null);
      fetchSupplierRFQs();
    } catch (err) {
      console.error("Failed to commit bid parameters:", err);
      alert("Submission error. Check system connectivity.");
    } finally {
      setIsUpdating(false);
    }
  };

  // Automated Browser Native CSV/Excel Exporter 
  const exportToExcel = () => {
    if (rfqs.length === 0) return;

    // Create CSV rows matching standard supply tracking sheets
    const headers = ["Item Number", "Description", "Qty Required", "UOM", "Status", "Your Price ($)", "Lead Time", "Notes"];
    const rows = rfqs.map(item => [
      `"${item.itemNumber}"`,
      `"${item.description.replace(/"/g, '""')}"`,
      item.quantity,
      `"${item.uom}"`,
      `"${item.status}"`,
      item.offeredPrice || "—",
      `"${item.leadTime || "—"}"`,
      `"${item.supplierNote.replace(/"/g, '""')}"`
    ]);

    const csvContent = "data:text/csv;charset=utf-8," 
      + [headers.join(","), ...rows.map(e => e.join(","))].join("\n");
    
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `RFQ_Quote_Sheet_${profile?.companyName || "Vendor"}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  if (loading) return <div className="p-8">Authenticating portal access profiles...</div>;

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <header className="mb-8 flex justify-between items-center border-b border-slate-200 pb-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">Vendor Bidding Terminal</h1>
          <p className="text-sm text-slate-500">
            Account Workspace: <span className="font-semibold text-slate-700">{profile?.companyName || "Authorized Supplier"}</span>
          </p>
        </div>
        <div>
          <button
            onClick={exportToExcel}
            disabled={rfqs.length === 0}
            className="rounded-md bg-emerald-600 px-4 py-2 text-xs font-semibold text-white shadow-sm hover:bg-emerald-500 transition-colors disabled:opacity-50"
          >
            Export Sheet to Excel (.CSV)
          </button>
        </div>
      </header>

      {/* Main Action Workspace Grid */}
      <div className="bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-200">
          <h3 className="text-base font-semibold text-slate-900">Your Open Sourcing Requests</h3>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-sm">
            <thead className="bg-slate-100 text-slate-700 uppercase tracking-wider text-xs border-b border-slate-200">
              <tr>
                <th className="py-3 px-6">Item #</th>
                <th className="py-3 px-6">Description</th>
                <th className="py-3 px-6 text-right">Qty</th>
                <th className="py-3 px-6">UOM</th>
                <th className="py-3 px-6">Unit Price ($)</th>
                <th className="py-3 px-6">Est. Lead Time</th>
                <th className="py-3 px-6">Supplier Remarks</th>
                <th className="py-3 px-6 text-center">Execution</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 text-slate-800">
              {rfqs.length === 0 ? (
                <tr>
                  <td colSpan={8} className="py-8 text-center text-slate-400">
                    No sourcing lines assigned to your profile currently.
                  </td>
                </tr>
              ) : (
                rfqs.map((item) => (
                  <tr key={item.id} className="hover:bg-slate-50 transition-colors">
                    <td className="py-4 px-6 font-mono font-medium text-slate-900">{item.itemNumber}</td>
                    <td className="py-4 px-6 max-w-xs truncate">{item.description}</td>
                    <td className="py-4 px-6 text-right font-medium">{item.quantity}</td>
                    <td className="py-4 px-6 text-slate-500">{item.uom}</td>
                    
                    {editingId === item.id ? (
                      <>
                        <td className="py-4 px-2">
                          <input
                            type="number"
                            step="0.01"
                            value={price}
                            onChange={(e) => setPrice(e.target.value)}
                            className="w-24 rounded border border-slate-300 px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-blue-500"
                            placeholder="0.00"
                          />
                        </td>
                        <td className="py-4 px-2">
                          <input
                            type="text"
                            value={leadTime}
                            onChange={(e) => setLeadTime(e.target.value)}
                            className="w-28 rounded border border-slate-300 px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-blue-500"
                            placeholder="e.g., 2 weeks"
                          />
                        </td>
                        <td className="py-4 px-2">
                          <input
                            type="text"
                            value={note}
                            onChange={(e) => setNote(e.target.value)}
                            className="w-full min-w-[150px] rounded border border-slate-300 px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-blue-500"
                            placeholder="Add item conditions..."
                          />
                        </td>
                        <td className="py-4 px-6 text-center whitespace-nowrap">
                          <button
                            onClick={() => handleUpdateBid(item.id)}
                            disabled={isUpdating}
                            className="rounded bg-blue-600 px-2.5 py-1 text-xs font-semibold text-white shadow-sm hover:bg-blue-500 mr-2"
                          >
                            Save
                          </button>
                          <button
                            onClick={() => setEditingId(null)}
                            className="rounded bg-slate-200 px-2.5 py-1 text-xs font-semibold text-slate-700 hover:bg-slate-300"
                          >
                            Cancel
                          </button>
                        </td>
                      </>
                    ) : (
                      <>
                        <td className="py-4 px-6 font-semibold text-slate-900">
                          {item.offeredPrice !== null ? `$${item.offeredPrice.toFixed(2)}` : <span className="text-slate-300">——</span>}
                        </td>
                        <td className="py-4 px-6 text-slate-700">{item.leadTime || <span className="text-slate-300">——</span>}</td>
                        <td className="py-4 px-6 text-xs text-slate-500 max-w-xs truncate" title={item.supplierNote}>
                          {item.supplierNote || <span className="text-slate-300">No notes recorded</span>}
                        </td>
                        <td className="py-4 px-6 text-center">
                          <button
                            onClick={() => startEditing(item)}
                            className="rounded border border-slate-300 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 transition-colors"
                          >
                            Quote Item
                          </button>
                        </td>
                      </>
                    )}
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}