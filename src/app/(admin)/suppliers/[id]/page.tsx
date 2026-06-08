// src/app/(admin)/suppliers/[id]/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import { collection, getDocs, query, where } from "firebase/firestore";
import { db } from "@/lib/firebase";
import { useAuth } from "@/context/AuthContext";
import { useRouter, useParams } from "next/navigation";

interface RoutedRFQ {
  id: string;
  itemNumber: string;
  description: string;
  quantity: number;
  uom: string;
  status: string;
  offeredPrice: number | null;
  leadTime: string | null;
  supplierNote: string;
}

export default function SupplierMirrorView() {
  const { profile, loading } = useAuth();
  const router = useRouter();
  const params = useParams();
  
  const [mirroredRfqs, setMirroredRfqs] = useState<RoutedRFQ[]>([]);
  const [isMirrorLoading, setIsMirrorLoading] = useState(true);
  const [companyName, setCompanyName] = useState<string>("Supplier");

  useEffect(() => {
    if (!loading && (!profile || profile.role !== "admin")) {
      router.push("/login");
    }
  }, [profile, loading, router]);

  useEffect(() => {
    if (profile?.role === "admin" && params?.id) {
      fetchMirroredVendorView(params.id as string);
    }
  }, [profile, params]);

  const fetchMirroredVendorView = async (supplierUid: string) => {
    setIsMirrorLoading(true);
    try {
      // Pull items matching this supplier's authentication UID
      const q = query(collection(db, "rfq_routing"), where("supplierId", "==", supplierUid));
      const snapshot = await getDocs(q);
      const list: RoutedRFQ[] = [];
      
      snapshot.forEach((doc) => {
        list.push({ id: doc.id, ...doc.data() } as RoutedRFQ);
      });
      setMirroredRfqs(list);

      // Extract company name if items exist
      if (snapshot.docs.length > 0) {
        // Fallback or use a separate quick fetch for company metadata if needed
        const data = snapshot.docs[0].data();
        // Since we store company name context, we can extract or fallback gracefully
      }
    } catch (err) {
      console.error("Error loading mirror queue:", err);
    } finally {
      setIsMirrorLoading(false);
    }
  };

  if (loading) return <div className="p-8">Verifying credentials...</div>;

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <header className="mb-8 flex justify-between items-center border-b border-slate-200 pb-4">
        <div>
          <button 
            onClick={() => router.push("/suppliers")}
            className="text-xs font-bold text-slate-500 hover:text-slate-700 uppercase tracking-wider mb-2 block"
          >
            ← Back to Suppliers Directory
          </button>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">Administrative Mirror Workspace</h1>
          <p className="text-sm text-slate-500">Simulating live, read-only interface views for assigned vendor profiles</p>
        </div>
        <span className="inline-flex items-center rounded-md bg-amber-50 px-2.5 py-1 text-xs font-semibold text-amber-800 ring-1 ring-inset ring-amber-600/20">
          Admin Audit Active
        </span>
      </header>

      <div className="bg-white border border-blue-200 rounded-lg shadow-sm overflow-hidden ring-1 ring-blue-500/10">
        <div className="px-6 py-4 border-b border-blue-200 bg-blue-50/50">
          <h3 className="text-sm font-bold uppercase text-blue-900 tracking-wider">
            Live Open Sourcing Requests Queue
          </h3>
        </div>
        
        <div className="overflow-x-auto">
          {isMirrorLoading ? (
            <div className="p-12 text-center text-slate-500">Querying live supplier queue database...</div>
          ) : (
            <table className="w-full text-left border-collapse text-sm">
              <thead className="bg-slate-50 text-slate-700 font-semibold text-xs border-b border-slate-200">
                <tr>
                  <th className="py-3 px-6">Item #</th>
                  <th className="py-3 px-6">Description</th>
                  <th className="py-3 px-6 text-right">Qty Required</th>
                  <th className="py-3 px-6">UOM</th>
                  <th className="py-3 px-6">Current Price ($)</th>
                  <th className="py-3 px-6">Stated Lead Time</th>
                  <th className="py-3 px-6">Supplier Notes</th>
                  <th className="py-3 px-6 text-center">Bidding Status</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 text-slate-800">
                {mirroredRfqs.length === 0 ? (
                  <tr>
                    <td colSpan={8} className="py-12 text-center text-slate-400">
                      This supplier has not been routed any material line items yet.
                    </td>
                  </tr>
                ) : (
                  mirroredRfqs.map((item) => (
                    <tr key={item.id} className="hover:bg-slate-50/50">
                      <td className="py-4 px-6 font-mono font-medium text-slate-900">{item.itemNumber}</td>
                      <td className="py-4 px-6 max-w-xs truncate">{item.description}</td>
                      <td className="py-4 px-6 text-right font-medium">{item.quantity}</td>
                      <td className="py-4 px-6 text-slate-500">{item.uom}</td>
                      <td className="py-4 px-6 font-semibold text-slate-900">
                        {item.offeredPrice !== null ? `$${item.offeredPrice.toFixed(2)}` : <span className="text-slate-300">——</span>}
                      </td>
                      <td className="py-4 px-6 text-slate-700">{item.leadTime || <span className="text-slate-300">——</span>}</td>
                      <td className="py-4 px-6 text-xs text-slate-500 max-w-xs truncate" title={item.supplierNote}>
                        {item.supplierNote || <span className="text-slate-300">—</span>}
                      </td>
                      <td className="py-4 px-6 text-center">
                        <span className={`inline-flex items-center rounded-full px-2.5 py-0.5 text-xs font-medium ${
                          item.status === 'Pending' ? 'bg-amber-50 text-amber-800 ring-1 ring-amber-600/20' : 'bg-emerald-50 text-emerald-800 ring-1 ring-emerald-600/20'
                        }`}>
                          {item.status}
                        </span>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  );
}