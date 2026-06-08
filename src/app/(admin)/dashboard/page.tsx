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
  supplierId: string;
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

  // Modal State Control
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [selectedItem, setSelectedItem] = useState<MaterialItem | null>(null);
  const [selectedSupplierIds, setSelectedSupplierIds] = useState<string[]>([]);
  const [isRouting, setIsRouting] = useState(false);

  // Router Authorization Check
  useEffect(() => {
    if (!loading && (!profile || profile.role !== "admin")) {
      router.push("/login");
    }
  }, [profile, loading, router]);

  // Load Core Collections
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
    setSelectedSupplierIds([]);
    setIsModalOpen(true);
  };

  const handleToggleSupplierSelection = (supplierId: string) => {
    setSelectedSupplierIds((prev) =>
      prev.includes(supplierId)
        ? prev.filter((id) => id !== supplierId)
        : [...prev, supplierId]
    );
  };

  const handleDispatchRFQ = async () => {
    if (!selectedItem || selectedSupplierIds.length === 0) return;
    setIsRouting(true);

    try {
      const batch = writeBatch(db);

      selectedSupplierIds.forEach((supId) => {
        const rfqDocRef = doc(collection(db, "rfq_routing"));
        batch.set(rfqDocRef, {
          materialId: selectedItem.id,
          itemNumber: selectedItem.itemNumber,
          description: selectedItem.description,
          quantity: selectedItem.quantity,
          uom: selectedItem.uom,
          supplierId: supId, // Maps to specific Supplier Account UID
          status: "Pending",
          offeredPrice: null,
          leadTime: null,
          supplierNote: "",
          timestamp: new Date(),
        });
      });

      // Update parent status of item
      const materialDocRef = doc(db, "materials", selectedItem.id);
      batch.update(materialDocRef, { status: "Sourced" });

      await batch.commit();
      
      setIsModalOpen(false);
      fetchMaterials();
    } catch (err) {
      console.error("RFQ routing dispatch failed:", err);
      alert("System routing failure. Check network logs.");
    } finally {
      setIsRouting(false);
    }
  };

  if (loading) return <div className="p-8">Verifying credentials...</div>;

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <header className="mb-8 flex justify-between items-center border-b border-slate-200 pb-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">Material Procurement Console</h1>
          <p className="text-sm text-slate-500">Upload bulk material logs and dispatch item sourcing requirements</p>
        </div>
        <div className="flex items-center gap-4">
          <button
            onClick={() => router.push("/suppliers")}
            className="text-sm font-semibold text-blue-600 hover:text-blue-800 bg-blue-50 px-3 py-1.5 rounded-md transition-colors"
          >
            Manage Suppliers →
          </button>
          <span className="inline-flex items-center rounded-md bg-blue-50 px-2.5 py-0.5 text-xs font-medium text-blue-700 ring-1 ring-inset ring-blue-700/10">
            System Administrator
          </span>
        </div>
      </header>

      {/* Control Pane */}
      <div className="mb-8 p-6 bg-white border border-slate-200 rounded-lg shadow-sm">
        <h2 className="text-sm font-semibold text-slate-900 mb-2">Import Bulk Materials List</h2>
        <p className="text-xs text-slate-500 mb-4">Supported format: CSV with columns labeled [Item, Description, Qty, UOM]</p>
        
        <div className="flex items-center gap-4">
          <input
            type="file"
            accept=".csv"
            onChange={handleCSVUpload}
            disabled={isUploading}
            className="text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 cursor-pointer disabled:opacity-50"
          />
          {feedbackMessage && (
            <p className="text-sm font-medium text-slate-700 bg-slate-100 px-3 py-1 rounded">
              {feedbackMessage}
            </p>
          )}
        </div>
      </div>

      {/* Data Table View */}
      <div className="bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-200">
          <h3 className="text-base font-semibold text-slate-900">Master Line Items</h3>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-sm">
            <thead className="bg-slate-100 text-slate-700 uppercase tracking-wider text-xs border-b border-slate-200">
              <tr>
                <th className="py-3 px-6">Item #</th>
                <th className="py-3 px-6">Description</th>
                <th className="py-3 px-6 text-right">Qty Required</th>
                <th className="py-3 px-6">UOM</th>
                <th className="py-3 px-6 text-center">Status</th>
                <th className="py-3 px-6 text-center">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 text-slate-800">
              {materials.length === 0 ? (
                <tr>
                  <td colSpan={6} className="py-8 text-center text-slate-400">
                    No records imported yet.
                  </td>
                </tr>
              ) : (
                materials.map((item) => (
                  <tr key={item.id} className="hover:bg-slate-50 transition-colors">
                    <td className="py-4 px-6 font-mono font-medium text-slate-900">{item.itemNumber}</td>
                    <td className="py-4 px-6 max-w-md truncate">{item.description}</td>
                    <td className="py-4 px-6 text-right font-semibold">{item.quantity}</td>
                    <td className="py-4 px-6 text-slate-500">{item.uom}</td>
                    <td className="py-4 px-6 text-center">
                      <span className={`inline-flex items-center rounded-full px-2.5 py-0.5 text-xs font-medium ring-1 ring-inset ${
                        item.status === 'Pending' ? 'bg-yellow-50 text-yellow-800 ring-yellow-600/20' : 'bg-blue-50 text-blue-800 ring-blue-600/20'
                      }`}>
                        {item.status}
                      </span>
                    </td>
                    <td className="py-4 px-6 text-center">
                      <button 
                        onClick={() => openSourcingModal(item)}
                        className="rounded bg-blue-600 px-3 py-1.5 text-xs font-semibold text-white shadow-sm hover:bg-blue-500 transition-colors"
                      >
                        Source Item
                      </button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Sourcing Dispatch Modal */}
      {isModalOpen && selectedItem && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="w-full max-w-lg rounded-lg bg-white p-6 shadow-xl border border-slate-200 animate-in fade-in zoom-in-95 duration-150">
            <div className="border-b border-slate-200 pb-3 mb-4">
              <h3 className="text-lg font-bold text-slate-900">Select Procurement Sources</h3>
              <p className="text-xs text-slate-500 mt-1">Item: <span className="font-mono font-bold text-slate-700">{selectedItem.itemNumber}</span> - {selectedItem.description}</p>
            </div>

            <div className="max-h-60 overflow-y-auto mb-6 space-y-2 pr-1">
              {suppliers.length === 0 ? (
                <p className="text-sm text-slate-400 py-4 text-center">No active suppliers registered. Set up profiles in the Supplier Directory.</p>
              ) : (
                suppliers.map((supplier) => (
                  <label 
                    key={supplier.id} 
                    className={`flex items-center justify-between p-3 rounded-md border text-sm cursor-pointer transition-all ${
                      selectedSupplierIds.includes(supplier.supplierId)
                        ? 'border-blue-500 bg-blue-50/50'
                        : 'border-slate-200 hover:bg-slate-50'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <input
                        type="checkbox"
                        checked={selectedSupplierIds.includes(supplier.supplierId)}
                        onChange={() => handleToggleSupplierSelection(supplier.supplierId)}
                        className="h-4 w-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500"
                      />
                      <div>
                        <p className="font-semibold text-slate-900">{supplier.companyName}</p>
                        <p className="text-xs text-slate-500">{supplier.email}</p>
                      </div>
                    </div>
                  </label>
                ))
              )}
            </div>

            <div className="flex justify-end gap-3 border-t border-slate-200 pt-4">
              <button
                type="button"
                onClick={() => setIsModalOpen(false)}
                className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50 transition-colors"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleDispatchRFQ}
                disabled={selectedSupplierIds.length === 0 || isRouting}
                className="rounded-md bg-blue-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-blue-500 disabled:bg-blue-300 transition-colors"
              >
                {isRouting ? "Routing RFQs..." : `Dispatch to ${selectedSupplierIds.length} Vendor(s)`}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}