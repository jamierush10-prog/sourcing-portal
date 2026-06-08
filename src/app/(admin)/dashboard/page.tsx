// src/app/(admin)/dashboard/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import Papa from "papaparse";
import { collection, writeBatch, doc, getDocs, query, orderBy } from "firebase/firestore";
import { db } from "@/lib/firebase";
import { useAuth } from "@/context/AuthContext";
import { useRouter } from "next/navigation";

interface MaterialItem {
  id: string;
  itemNumber: string;
  description: string;
  quantity: number;
  uom: string;
  status: "Pending" | "Sourced" | "Completed";
}

interface SupplierProfile {
  id: string;
  supplierNo: string;
  companyName: string;
  contactName: string;
  email: string;
}

export default function AdminDashboard() {
  const { profile, loading } = useAuth();
  const router = useRouter();
  
  const [materials, setMaterials] = useState<MaterialItem[]>([]);
  const [suppliers, setSuppliers] = useState<SupplierProfile[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [feedbackMessage, setFeedbackMessage] = useState("");

  // Sourcing Modal States
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [selectedItem, setSelectedItem] = useState<MaterialItem | null>(null);
  const [selectedSupplierNos, setSelectedSupplierNos] = useState<string[]>([]);
  const [isRouting, setIsRouting] = useState(false);

  useEffect(() => {
    if (!loading && (!profile || profile.role !== "admin")) {
      router.push("/login");
    }
  }, [profile, loading, router]);

  useEffect(() => {
    if (profile?.role === "admin") {
      fetchMaterials();
      fetchSuppliers();
    }
  }, [profile]);

  const fetchMaterials = async () => {
    try {
      const q = query(collection(db, "materials"), orderBy("itemNumber", "asc"));
      const querySnapshot = await getDocs(q);
      const items: MaterialItem[] = [];
      querySnapshot.forEach((doc) => {
        items.push({ id: doc.id, ...doc.data() } as MaterialItem);
      });
      setMaterials(items);
    } catch (err) {
      console.error("Error fetching materials: ", err);
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
      console.error("Error fetching suppliers:", err);
    }
  };

  const handleCSVUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsUploading(true);
    setFeedbackMessage("Parsing file data...");

    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: async (results) => {
        try {
          const batch = writeBatch(db);
          
          results.data.forEach((row: any) => {
            const newDocRef = doc(collection(db, "materials"));
            batch.set(newDocRef, {
              itemNumber: row["Item"] || row["itemNumber"] || "UNKNOWN",
              description: row["Description"] || row["description"] || "",
              quantity: Number(row["Qty"] || row["quantity"] || 0),
              uom: row["UOM"] || row["uom"] || "EA",
              status: "Pending",
            });
          });

          setFeedbackMessage("Committing records to Firestore...");
          await batch.commit();
          setFeedbackMessage("Material list successfully updated!");
          fetchMaterials();
        } catch (error) {
          console.error("Batch write failure:", error);
          setFeedbackMessage("Upload failed. Verify column structures.");
        } finally {
          setIsUploading(false);
        }
      },
    });
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
          itemNumber: selectedItem.itemNumber,
          description: selectedItem.description,
          quantity: selectedItem.quantity,
          uom: selectedItem.uom,
          supplierNo: supNo, // Stamped directly with short code key
          status: "Pending",
          offeredPrice: null,
          leadTime: null,
          supplierNote: "",
          timestamp: new Date(),
        });
      });

      const materialDocRef = doc(db, "materials", selectedItem.id);
      batch.update(materialDocRef, { status: "Sourced" });

      await batch.commit();
      setIsModalOpen(false);
      fetchMaterials();
    } catch (err) {
      console.error("RFQ routing dispatch failed:", err);
    } finally {
      setIsRouting(false);
    }
  };

  if (loading) return <div className="p-8 text-sm text-slate-500">Verifying credentials...</div>;

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <header className="mb-8 flex justify-between items-center border-b border-slate-200 pb-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">Material Procurement Console</h1>
          <p className="text-sm text-slate-500">Parse master material parameters and assign vendor routing configurations</p>
        </div>
        <div className="flex items-center gap-4">
          <button
            onClick={() => router.push("/suppliers")}
            className="text-sm font-semibold text-blue-600 hover:text-blue-800 bg-blue-50 px-3 py-1.5 rounded-md transition-colors"
          >
            Manage Suppliers Directory →
          </button>
          <span className="inline-flex items-center rounded-md bg-blue-50 px-2.5 py-0.5 text-xs font-medium text-blue-700 ring-1 ring-inset ring-blue-700/10">
            System Administrator
          </span>
        </div>
      </header>

      {/* CSV Import Section */}
      <div className="mb-8 p-4 bg-white border border-slate-200 rounded-lg shadow-sm flex flex-col sm:flex-row gap-4 justify-between items-start sm:items-center">
        <div>
          <h4 className="text-sm font-bold text-slate-800">Import Master Material Requirements</h4>
          <p className="text-xs text-slate-500">Upload bulk CSV listings targeting Item, Description, Qty, and UOM parameters</p>
        </div>
        <div className="flex items-center gap-2 w-full sm:w-auto">
          <input
            type="file"
            accept=".csv"
            onChange={handleCSVUpload}
            disabled={isUploading}
            className="text-xs text-slate-500 file:mr-2 file:py-1 file:px-3 file:rounded file:border-0 file:text-xs file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 cursor-pointer"
          />
          {feedbackMessage && <span className="text-xs font-medium text-slate-600 bg-slate-100 px-2 py-1 rounded">{feedbackMessage}</span>}
        </div>
      </div>

      {/* Materials Log Grid */}
      <div className="bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-200 bg-slate-50/70">
          <h3 className="text-sm font-bold uppercase text-slate-700 tracking-wider">Master Requirements Queue</h3>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-sm">
            <thead className="bg-slate-100 text-slate-700 font-semibold text-xs border-b border-slate-200">
              <tr>
                <th className="py-3 px-6">Item #</th>
                <th className="py-3 px-6">Description</th>
                <th className="py-3 px-6 text-right">Qty</th>
                <th className="py-3 px-6">UOM</th>
                <th className="py-3 px-6 text-center">Status</th>
                <th className="py-3 px-6 text-center">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 text-slate-800">
              {materials.length === 0 ? (
                <tr><td colSpan={6} className="py-8 text-center text-slate-400">No records imported yet.</td></tr>
              ) : (
                materials.map((item) => (
                  <tr key={item.id} className="hover:bg-slate-50">
                    <td className="py-4 px-6 font-mono font-medium text-slate-900">{item.itemNumber}</td>
                    <td className="py-4 px-6 max-w-md truncate">{item.description}</td>
                    <td className="py-4 px-6 text-right font-semibold">{item.quantity}</td>
                    <td className="py-4 px-6 text-slate-500">{item.uom}</td>
                    <td className="py-4 px-6 text-center">
                      <span className={`inline-flex items-center rounded-full px-2 py-0.5 text-xs font-medium ${
                        item.status === 'Pending' ? 'bg-yellow-50 text-yellow-800 ring-1 ring-yellow-600/20' : 'bg-blue-50 text-blue-800 ring-1 ring-blue-600/20'
                      }`}>{item.status}</span>
                    </td>
                    <td className="py-4 px-6 text-center">
                      <button onClick={() => openSourcingModal(item)} className="rounded bg-blue-600 px-3 py-1.5 text-xs font-semibold text-white shadow-sm hover:bg-blue-500">Source Item</button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Sourcing Modal Display */}
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
              <button type="button" onClick={() => setIsModalOpen(false)} className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50">Cancel</button>
              <button type="button" onClick={handleDispatchRFQ} disabled={selectedSupplierNos.length === 0 || isRouting} className="rounded-md bg-blue-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-blue-500 disabled:bg-blue-300">
                {isRouting ? "Routing RFQs..." : `Dispatch to ${selectedSupplierNos.length} Vendor(s)`}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}