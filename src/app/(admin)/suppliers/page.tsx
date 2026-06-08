// src/app/(admin)/suppliers/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import { collection, addDoc, doc, updateDoc, getDocs, query, orderBy } from "firebase/firestore";
import { db } from "@/lib/firebase";
import { useAuth } from "@/context/AuthContext";
import { useRouter } from "next/navigation";

interface SupplierProfile {
  id: string;
  supplierNo: string;
  companyName: string;
  contactName: string;
  email: string;
}

export default function SuppliersDirectory() {
  const { profile, loading } = useAuth();
  const router = useRouter();

  const [suppliers, setSuppliers] = useState<SupplierProfile[]>([]);
  
  // Tracking form states
  const [editingSupplier, setEditingSupplier] = useState<SupplierProfile | null>(null);
  const [companyName, setCompanyName] = useState("");
  const [contactName, setContactName] = useState("");
  const [email, setEmail] = useState("");
  const [supplierNo, setSupplierNo] = useState(""); 
  
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");

  useEffect(() => {
    if (!loading && (!profile || profile.role !== "admin")) {
      router.push("/login");
    }
  }, [profile, loading, router]);

  useEffect(() => {
    if (profile?.role === "admin") {
      fetchSuppliers();
    }
  }, [profile]);

  const fetchSuppliers = async () => {
    try {
      const q = query(collection(db, "suppliers"), orderBy("companyName", "asc"));
      const querySnapshot = await getDocs(q);
      const list: SupplierProfile[] = [];
      querySnapshot.forEach((doc) => {
        list.push({ id: doc.id, ...doc.data() } as SupplierProfile);
      });
      setSuppliers(list);
    } catch (err) {
      console.error("Error pulling registry documents:", err);
    }
  };

  const handleStartEdit = (supplier: SupplierProfile) => {
    setErrorMessage("");
    setEditingSupplier(supplier);
    setCompanyName(supplier.companyName);
    setContactName(supplier.contactName || "");
    setEmail(supplier.email);
    setSupplierNo(supplier.supplierNo);
  };

  const handleCancelEdit = () => {
    setEditingSupplier(null);
    setCompanyName("");
    setContactName("");
    setEmail("");
    setSupplierNo("");
    setErrorMessage("");
  };

  const handleSaveSupplierProfile = async (e: React.FormEvent) => {
    e.preventDefault();
    setErrorMessage("");
    
    const formattedNo = supplierNo.trim().toUpperCase();
    
    // Validate 4-6 character constraints
    if (formattedNo.length < 4 || formattedNo.length > 6) {
      setErrorMessage("Supplier Number must be between 4 and 6 characters.");
      return;
    }

    // Check for duplicate Supplier Numbers among OTHER companies
    const isDuplicate = suppliers.some(
      (s) => s.supplierNo === formattedNo && s.id !== editingSupplier?.id
    );
    if (isDuplicate) {
      setErrorMessage(`Supplier Number "${formattedNo}" is already assigned to another vendor.`);
      return;
    }

    setIsSubmitting(true);
    try {
      const payload = {
        companyName: companyName.trim(),
        contactName: contactName.trim(),
        email: email.trim(),
        supplierNo: formattedNo,
      };

      if (editingSupplier) {
        // Update existing document inside suppliers collection
        const supplierDocRef = doc(db, "suppliers", editingSupplier.id);
        await updateDoc(supplierDocRef, payload);
      } else {
        // Create brand new entry link
        await addDoc(collection(db, "suppliers"), payload);
      }

      handleCancelEdit();
      fetchSuppliers();
    } catch (err) {
      console.error("Failed to commit vendor registry profile modification:", err);
      setErrorMessage("Error saving record parameters. Please try again.");
    } finally {
      setIsSubmitting(false);
    }
  };

  if (loading) return <div className="p-8 text-sm text-slate-500">Verifying credentials...</div>;

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <header className="mb-8 flex justify-between items-center border-b border-slate-200 pb-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">Supplier Directory Configuration</h1>
          <p className="text-sm text-slate-500">Manage active company short codes and adjust master distribution parameters</p>
        </div>
        <button
          onClick={() => router.push("/dashboard")}
          className="text-sm font-semibold text-blue-600 hover:text-blue-800 bg-blue-50 px-3 py-1.5 rounded-md transition-colors"
        >
          ← Return to Console
        </button>
      </header>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
        {/* Dynamic Vendor Form Card (Handles Create or Edit States) */}
        <div className="bg-white p-6 border border-slate-200 rounded-lg shadow-sm h-fit">
          <div className="mb-4">
            <h3 className="text-sm font-bold uppercase text-slate-700 tracking-wider">
              {editingSupplier ? "⚙ Modify Vendor Registry" : "Link Vendor Profile"}
            </h3>
            {editingSupplier && (
              <p className="text-xs text-slate-400 mt-0.5">Updating tracking record reference ID: {editingSupplier.id}</p>
            )}
          </div>

          <form onSubmit={handleSaveSupplierProfile} className="space-y-4">
            {errorMessage && (
              <div className="p-3 text-xs font-semibold text-red-800 bg-red-50 border border-red-200 rounded">
                {errorMessage}
              </div>
            )}
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">Supplier No * (4-6 Chars)</label>
              <input
                type="text"
                required
                maxLength={6}
                value={supplierNo}
                onChange={(e) => setSupplierNo(e.target.value)}
                className="w-full text-sm rounded border border-slate-300 px-3 py-2 font-mono font-bold uppercase focus:outline-none focus:ring-1 focus:ring-blue-500 bg-slate-50"
                placeholder="e.g. PSTN"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">Company Name *</label>
              <input
                type="text"
                required
                value={companyName}
                onChange={(e) => setCompanyName(e.target.value)}
                className="w-full text-sm rounded border border-slate-300 px-3 py-2 focus:outline-none focus:ring-1 focus:ring-blue-500"
                placeholder="Piston Supply Co."
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">Primary Contact Name</label>
              <input
                type="text"
                value={contactName}
                onChange={(e) => setContactName(e.target.value)}
                className="w-full text-sm rounded border border-slate-300 px-3 py-2 focus:outline-none focus:ring-1 focus:ring-blue-500"
                placeholder="John Doe"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">Email Address *</label>
              <input
                type="email"
                required
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                className="w-full text-sm rounded border border-slate-300 px-3 py-2 focus:outline-none focus:ring-1 focus:ring-blue-500"
                placeholder="sales@piston.com"
              />
            </div>

            <div className="pt-2 space-y-2">
              <button
                type="submit"
                disabled={isSubmitting}
                className={`w-full rounded py-2 text-sm font-semibold text-white shadow-sm transition-colors ${
                  editingSupplier ? 'bg-amber-600 hover:bg-amber-500' : 'bg-blue-600 hover:bg-blue-500'
                } disabled:bg-slate-300`}
              >
                {isSubmitting ? "Processing Database Sync..." : editingSupplier ? "Update Corporate Profile" : "Register Vendor Link"}
              </button>
              
              {editingSupplier && (
                <button
                  type="button"
                  onClick={handleCancelEdit}
                  className="w-full rounded border border-slate-300 bg-white py-2 text-sm font-medium text-slate-700 hover:bg-slate-50 transition-colors"
                >
                  Cancel Parameter Changes
                </button>
              )}
            </div>
          </form>
        </div>

        {/* Master Registered Suppliers Directory Table Grid */}
        <div className="lg:col-span-2 bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden">
          <div className="px-6 py-4 border-b border-slate-200 bg-slate-50/70">
            <h3 className="text-sm font-bold uppercase text-slate-700 tracking-wider">Authorized Company Accounts</h3>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-sm">
              <thead className="bg-slate-100 text-slate-700 font-semibold text-xs border-b border-slate-200">
                <tr>
                  <th className="py-3 px-6">Supplier No</th>
                  <th className="py-3 px-6">Company</th>
                  <th className="py-3 px-6">Contact</th>
                  <th className="py-3 px-6">Email</th>
                  <th className="py-3 px-6 text-center">System Management Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 text-slate-800">
                {suppliers.length === 0 ? (
                  <tr><td colSpan={5} className="py-8 text-center text-slate-400">No suppliers registered in the database registry yet.</td></tr>
                ) : (
                  suppliers.map((supplier) => (
                    <tr key={supplier.id} className="hover:bg-slate-50">
                      <td className="py-4 px-6 font-mono font-bold text-blue-600 bg-blue-50/30">{supplier.supplierNo}</td>
                      <td className="py-4 px-6 font-semibold text-slate-900">{supplier.companyName}</td>
                      <td className="py-4 px-6 text-slate-600">{supplier.contactName || "—"}</td>
                      <td className="py-4 px-6 text-slate-500 font-mono text-xs">{supplier.email}</td>
                      <td className="py-4 px-6 text-center whitespace-nowrap space-x-2">
                        <button
                          onClick={() => handleStartEdit(supplier)}
                          className="inline-flex items-center rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs font-semibold text-slate-700 shadow-sm hover:bg-slate-50 transition-all"
                        >
                          Edit
                        </button>
                        <button
                          onClick={() => router.push(`/suppliers/${supplier.supplierNo}`)}
                          className="inline-flex items-center rounded bg-blue-600 px-2.5 py-1.5 text-xs font-semibold text-white shadow-sm hover:bg-blue-500 transition-all"
                        >
                          👁 Mirror Workspace
                        </button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}