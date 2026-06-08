// src/app/(supplier)/supplier-dashboard/page.tsx
"use client";

import React, { useState, useEffect } from "react";
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
}

export default function SupplierDashboard() {
  const { profile, user, loading } = useAuth();
  const router = useRouter();

  const [rfqs, setRfqs] = useState<RFQItem[]>([]);
  const [isDataLoading, setIsDataLoading] = useState(false);

  // Inline Editing States
  const [editingId, setEditingId] = useState<string | null>(null);
  const [bidPrice, setBidPrice] = useState<string>("");
  const [leadTime, setLeadTime] = useState<string>("");
  const [vendorNotes, setVendorNotes] = useState<string>("");
  const [isSaving, setIsSaving] = useState(false);

  useEffect(() => {
    if (!loading && (!profile || profile.role !== "supplier")) {
      router.push("/login");
    }
  }, [profile, loading, router]);

  useEffect(() => {
    if (!loading && profile?.role === "supplier" && user?.uid) {
      fetchSupplierRFQs();
    }
  }, [profile, user, loading]);

  const fetchSupplierRFQs = async () => {
    setIsDataLoading(true);
    try {
      const q = query(
        collection(db, "rfq_routing"),
        where("supplierId", "==", user?.uid)
      );
      const snapshot = await getDocs(q);
      const list: RFQItem[] = [];
      snapshot.forEach((doc) => {
        list.push({ id: doc.id, ...doc.data() } as RFQItem);
      });
      setRfqs(list);
    } catch (err) {
      console.error("Error pulling isolated vendor payload:", err);
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
      
      // Update the routing entry with vendor metrics
      await updateDoc(rfqDocRef, {
        offeredPrice: parsedPrice,
        leadTime: leadTime.trim(),
        supplierNote: vendorNotes.trim(),
        status: "Completed" // Switches status to complete once they bid
      });

      setEditingId(null);
      fetchSupplierRFQs(); // Refresh the table metrics
    } catch (err) {
      console.error("Failed to commit supplier bid data:", err);
      alert("Error saving your bid. Please try again.");
    } finally {
      setIsSaving(false);
    }
  };

  if (loading) return <div className="p-8 text-sm text-slate-500">Verifying security parameters...</div>;

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <header className="mb-8 flex justify-between items-center border-b border-slate-200 pb-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">
            {profile?.companyName || "Vendor"} Bidding Terminal
          </h1>
          <p className="text-sm text-slate-500">Review open material requirements, log quote pricing, and specify execution lead times</p>
        </div>
        <span className="inline-flex items-center rounded-md bg-emerald-50 px-2.5 py-0.5 text-xs font-medium text-emerald-700 ring-1 ring-inset ring-emerald-700/10">
          Authorized Supplier Portal
        </span>
      </header>

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
                  <th className="py-3 px-6">Item #</th>
                  <th className="py-3 px-6">Description</th>
                  <th className="py-3 px-6 text-right">Qty</th>
                  <th className="py-3 px-6">UOM</th>
                  <th className="py-3 px-6">Your Price ($)</th>
                  <th className="py-3 px-6">Lead Time</th>
                  <th className="py-3 px-6">Notes</th>
                  <th className="py-3 px-6 text-center">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 text-slate-800">
                {rfqs.length === 0 ? (
                  <tr>
                    <td colSpan={8} className="py-12 text-center text-slate-400">
                      No material requests dispatched to your profile queue yet.
                    </td>
                  </tr>
                ) : (
                  rfqs.map((item) => {
                    const isEditing = editingId === item.id;
                    return (
                      <tr key={item.id} className={`hover:bg-slate-50/50 transition-colors ${isEditing ? 'bg-blue-50/30' : ''}`}>
                        <td className="py-4 px-6 font-mono font-medium text-slate-900">{item.itemNumber}</td>
                        <td className="py-4 px-6 max-w-xs truncate" title={item.description}>{item.description}</td>
                        <td className="py-4 px-6 text-right font-medium">{item.quantity}</td>
                        <td className="py-4 px-6 text-slate-500">{item.uom}</td>
                        
                        {/* Unit Price Field */}
                        <td className="py-3 px-4">
                          {isEditing ? (
                            <input
                              type="number"
                              step="0.01"
                              value={bidPrice}
                              onChange={(e) => setBidPrice(e.target.value)}
                              className="w-24 rounded border border-slate-300 px-2 py-1 text-sm font-medium focus:outline-none focus:ring-1 focus:ring-blue-500"
                              placeholder="0.00"
                            />
                          ) : item.offeredPrice !== null ? (
                            <span className="font-semibold text-slate-900">${item.offeredPrice.toFixed(2)}</span>
                          ) : (
                            <span className="text-slate-300 font-medium">Pending Entry</span>
                          )}
                        </td>

                        {/* Lead Time Field */}
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

                        {/* Supplier Notes Field */}
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

                        {/* Interactive Buttons */}
                        <td className="py-3 px-6 text-center whitespace-nowrap">
                          {isEditing ? (
                            <div className="flex items-center justify-center gap-2">
                              <button
                                onClick={() => handleSaveBid(item.id)}
                                disabled={isSaving}
                                className="rounded bg-emerald-600 px-2.5 py-1 text-xs font-semibold text-white shadow-sm hover:bg-emerald-500 disabled:bg-emerald-300"
                              >
                                {isSaving ? "Saving..." : "Save"}
                              </button>
                              <button
                                onClick={cancelEditing}
                                className="rounded border border-slate-300 bg-white px-2.5 py-1 text-xs font-medium text-slate-700 hover:bg-slate-50"
                              >
                                Cancel
                              </button>
                            </div>
                          ) : (
                            <button
                              onClick={() => startEditing(item)}
                              className="inline-flex items-center rounded bg-blue-50 px-3 py-1.5 text-xs font-semibold text-blue-700 shadow-sm hover:bg-blue-100 transition-colors"
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
    </div>
  );
}