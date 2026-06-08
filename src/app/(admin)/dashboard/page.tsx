// src/app/(admin)/dashboard/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import Papa from "papaparse";
import { collection, writeBatch, doc, getDocs, query, orderBy, where } from "firebase/firestore";
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

export default function AdminDashboard() {
  const { profile, loading } = useAuth();
  const router = useRouter();
  
  const [materials, setMaterials] = useState<MaterialItem[]>([]);
  const [suppliers, setSuppliers] = useState<SupplierProfile[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [feedbackMessage, setFeedbackMessage] = useState("");

  // Supplier Mirroring States
  const [selectedMirrorSupplier, setSelectedMirrorSupplier] = useState<string>("master");
  const [mirroredRfqs, setMirroredRfqs] = useState<RoutedRFQ[]>([]);
  const [isMirrorLoading, setIsMirrorLoading] = useState(false);

  // Sourcing Modal States
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [selectedItem, setSelectedItem] = useState<MaterialItem | null>(null);
  const [selectedSupplierIds, setSelectedSupplierIds] = useState<string[]>([]);
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

  // Triggered whenever the admin switches the vendor dropdown view
  useEffect(() => {
    if (selectedMirrorSupplier !== "master") {
      fetchMirroredVendorView(selectedMirrorSupplier);
    }
  }, [selectedMirrorSupplier]);

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

  const fetchMirroredVendorView = async (supplierUid: string) => {
    setIsMirrorLoading(true);
    try {
      const q = query(collection(db, "rfq_routing"), where("supplierId", "==", supplierUid));
      const snapshot = await getDocs(q);
      const list: RoutedRFQ[] = [];
      snapshot.forEach((doc) => {
        list.push({ id: doc.id, ...doc.data() } as RoutedRFQ);
      });
      setMirroredRfqs(list);
    } catch (err) {
      console.error("Error loading mirror queue:", err);
    } finally {
      setIsMirrorLoading(false);
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
          setSelectedMirrorSupplier("master"); // Reset view back to master grid
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
      prev.includes(supplierId) ? prev.filter((id) => id !== supplierId) : [...prev, supplierId]
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
          supplierId: supId,
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

  if (loading) return <div className="p-8">Verifying credentials...</div>;

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <header className="mb-8 flex justify-between items-center border-b border-slate-200 pb-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">Material Procurement Console</h1>
          <p className="text-sm text-slate-500">Monitor overall material logic or select an isolated vendor mirror below</p>
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

      {/* View Switcher Controls */}
      <div className="mb-8 flex flex-col sm:flex-row gap-4 justify-between items-start sm:items-center p-4 bg-white border border-slate-200 rounded-lg shadow-sm">
        <div>
          <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">Active Terminal View Profile</label>
          <select
            value={selectedMirrorSupplier}
            onChange={(e) => setSelectedMirrorSupplier(e.target.value)}
            className="rounded-md border border-slate-300 bg-slate-50 px-3 py-1.5 text-sm font-medium text-slate-800 focus:outline-none focus:ring-1 focus:ring-blue-500 cursor-pointer"
          >
            <option value="master">★ Master Material Management Queue (Admin View)</option>
            {suppliers.map((sup) => (
              <option key={sup.id} value={sup.supplierId}>
                👁 Mirror Workspace: {sup.companyName}
              </option>
            ))}
          </select>
        </div>

        {selectedMirrorSupplier === "master" && (
          <div className="w-full sm:w-auto">
            <span className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">Import Materials List</span>
            <div className="flex items-center gap-2">
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
        )}
      </div>

      {/* Dynamic Queue Display */}
      {selectedMirrorSupplier === "master" ? (
        /* MASTER ADMIN VIEW */
        <div className="bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden">
          <div className="px-6 py-4 border-b border-slate-200 bg-slate-50/70">
            <h3 className="text-sm font-bold uppercase text-slate-700 tracking-wider">Master Material Requirements</h3>
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
      ) : (
        /* MIRRORED SUPPLIER WORKSPACE VIEW */
        <div className="bg-white border border-blue-200 rounded-lg shadow-sm overflow-hidden ring-1 ring-blue-500/10">
          <div className="px-6 py-4 border-b border-blue-200 bg-blue-50/50 flex justify-between items-center">
            <h3 className="text-sm font-bold uppercase text-blue-900 tracking-wider">
              Auditing Workspace Profile: {suppliers.find(s => s.supplierId === selectedMirrorSupplier)?.companyName}
            </h3>
            <span className="inline-flex items-center rounded-md bg-amber-50 px-2 py-1 text-xs font-semibold text-amber-800 ring-1 ring-inset ring-amber-600/20">
              Read-Only Administrative Simulation
            </span>
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
                    <tr><td colSpan={8} className="py-12 text-center text-slate-400">This supplier has not been routed any line items yet.</td></tr>
                  ) : (
                    mirroredRfqs.map((item) => (
                      <tr key={item.id} className="hover:bg-slate-50/50">
                        <td className="py-4 px-6 font-mono font-medium text-slate-900">{item.itemNumber}</td>
                        <td className="py-4 px-6 max-w-xs truncate">{item.description}</td>
                        <td className="py-4 px-6 text-right">{item.quantity}</td>
                        <td className="py-4 px-6 text-slate-500">{item.uom}</td>
                        <td className="py-4 px-6 font-semibold text-slate-900">
                          {item.offeredPrice !== null ? `$${item.offeredPrice.toFixed(2)}` : <span className="text-slate-300">No entry</span>}
                        </td>
                        <td className="py-4 px-6 text-slate-700">{item.leadTime || <span className="text-slate-300">No entry</span>}</td>
                        <td className="py-4 px-6 text-xs text-slate-500 max-w-xs truncate">{item.supplierNote || <span className="text-slate-300">—</span>}</td>
                        <td className="py-4 px-6 text-center">
                          <span className={`inline-flex items-center rounded-full px-2 py-0.5 text-xs font-medium ${
                            item.status === 'Pending' ? 'bg-amber-50 text-amber-800 ring-1 ring-amber-600/20' : 'bg-emerald-50 text-emerald-800 ring-1 ring-emerald-600/20'
                          }`}>{item.status}</span>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            )}
          </div>
        </div>
      )}

      {/* Sourcing Dispatch Modal */}
      {isModalOpen && selectedItem && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="w-full max-w-lg rounded-lg bg-white p-6 shadow-xl border border-slate-200">
            <div className="border-b border-slate-200 pb-3 mb-4">
              <h3 className="text-lg font-bold text-slate-900">Select Procurement Sources</h3>
              <p className="text-xs text-slate-500 mt-1">Item: <span className="font-mono font-bold text-slate-700">{selectedItem.itemNumber}</span></p>
            </div>
            <div className="max-h-60 overflow-y-auto mb-6 space-y-2 pr-1">
              {suppliers.map((supplier) => (
                <label key={supplier.id} className={`flex items-center justify-between p-3 rounded-md border text-sm cursor-pointer ${selectedSupplierIds.includes(supplier.supplierId) ? 'border-blue-500 bg-blue-50/50' : 'border-slate-200 hover:bg-slate-50'}`}>
                  <div className="flex items-center gap-3">
                    <input type="checkbox" checked={selectedSupplierIds.includes(supplier.supplierId)} onChange={() => handleToggleSupplierSelection(supplier.supplierId)} className="h-4 w-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500" />
                    <div>
                      <p className="font-semibold text-slate-900">{supplier.companyName}</p>
                      <p className="text-xs text-slate-500">{supplier.email}</p>
                    </div>
                  </div>
                </label>
              ))}
            </div>
            <div className="flex justify-end gap-3 border-t border-slate-200 pt-4">
              <button type="button" onClick={() => setIsModalOpen(false)} className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50">Cancel</button>
              <button type="button" onClick={handleDispatchRFQ} disabled={selectedSupplierIds.length === 0 || isRouting} className="rounded-md bg-blue-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-blue-500 disabled:bg-blue-300">
                {isRouting ? "Routing RFQs..." : `Dispatch to ${selectedSupplierIds.length} Vendor(s)`}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}