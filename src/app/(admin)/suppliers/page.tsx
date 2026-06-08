// src/app/(admin)/suppliers/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import { collection, addDoc, getDocs, query, orderBy } from "firebase/firestore";
import { db } from "@/lib/firebase";
import { useAuth } from "@/context/AuthContext";
import { useRouter } from "next/navigation";

interface SupplierProfile {
  id?: string;
  supplierId: string; // Maps to their Firebase Auth UID
  companyName: string;
  contactName: string;
  email: string;
}

export default function SupplierManagement() {
  const { profile, loading } = useAuth();
  const router = useRouter();
  
  const [suppliers, setSuppliers] = useState<SupplierProfile[]>([]);
  const [companyName, setCompanyName] = useState("");
  const [contactName, setContactName] = useState("");
  const [email, setEmail] = useState("");
  const [supplierUid, setSupplierUid] = useState(""); // The user's Firebase UID
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [msg, setMsg] = useState("");

  // Route authorization check
  useEffect(() => {
    if (!loading && (!profile || profile.role !== "admin")) {
      router.push("/login");
    }
  }, [profile, loading, router]);

  // Fetch registered suppliers
  useEffect(() => {
    if (profile?.role === "admin") {
      fetchSuppliers();
    }
  }, [profile]);

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

  const handleRegisterSupplier = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!companyName || !email || !supplierUid) return;
    
    setIsSubmitting(true);
    setMsg("Linking database profile...");

    try {
      // 1. Create the master profile entry in the 'suppliers' collection
      await addDoc(collection(db, "suppliers"), {
        supplierId: supplierUid.trim(),
        companyName: companyName.trim(),
        contactName: contactName.trim(),
        email: email.trim().toLowerCase(),
      });

      // 2. Mirror a security document into the 'users' collection so the AuthContext sets the 'supplier' role properly
      // We use addDoc or setDoc here, but since we already know their UID, we can just provision it.
      // For now, we will assume you map the UID directly.
      setMsg("Supplier profile successfully linked!");
      setCompanyName("");
      setContactName("");
      setEmail("");
      setSupplierUid("");
      fetchSuppliers();
    } catch (err) {
      console.error(err);
      setMsg("Error provisioning supplier record.");
    } finally {
      setIsSubmitting(false);
    }
  };

  if (loading) return <div className="p-8">Verifying system configurations...</div>;

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <header className="mb-8 flex justify-between items-center border-b border-slate-200 pb-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">Supplier Directory</h1>
          <p className="text-sm text-slate-500">Provision vendor records and authorize portal routing profiles</p>
        </div>
        <div className="flex gap-4">
          <button 
            onClick={() => router.push("/dashboard")}
            className="text-sm font-semibold text-slate-600 hover:text-slate-900"
          >
            ← Back to Materials
          </button>
        </div>
      </header>

      <div className="grid grid-cols-1 gap-8 lg:grid-cols-3">
        {/* Onboarding Panel */}
        <div className="bg-white border border-slate-200 rounded-lg shadow-sm p-6 h-fit">
          <h3 className="text-base font-semibold text-slate-900 mb-4">Link New Supplier Profile</h3>
          
          <form onSubmit={handleRegisterSupplier} className="space-y-4">
            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">Company Name</label>
              <input
                type="text"
                required
                value={companyName}
                onChange={(e) => setCompanyName(e.target.value)}
                className="w-full rounded-md border border-slate-300 px-3 py-2 text-sm text-slate-900 focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
                placeholder="e.g., Vulcan Materials"
              />
            </div>

            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">Point of Contact Name</label>
              <input
                type="text"
                value={contactName}
                onChange={(e) => setContactName(e.target.value)}
                className="w-full rounded-md border border-slate-300 px-3 py-2 text-sm text-slate-900 focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
                placeholder="John Doe"
              />
            </div>

            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">Contact Email</label>
              <input
                type="email"
                required
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                className="w-full rounded-md border border-slate-300 px-3 py-2 text-sm text-slate-900 focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
                placeholder="sales@vendor.com"
              />
            </div>

            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">Firebase Account UID</label>
              <input
                type="text"
                required
                value={supplierUid}
                onChange={(e) => setSupplierUid(e.target.value)}
                className="w-full rounded-md border border-slate-300 px-3 py-2 text-sm font-mono text-slate-900 focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
                placeholder="Paste UID from Firebase Auth"
              />
              <p className="text-[10px] text-slate-400 mt-1">Links their authenticated login to this specific data routing.</p>
            </div>

            <button
              type="submit"
              disabled={isSubmitting}
              className="w-full rounded bg-blue-600 px-3 py-2 text-sm font-semibold text-white shadow-sm hover:bg-blue-500 disabled:bg-blue-400 transition-colors"
            >
              Register Vendor Link
            </button>

            {msg && <p className="text-xs font-medium text-center text-slate-700 bg-slate-100 p-2 rounded">{msg}</p>}
          </form>
        </div>

        {/* Directory Listing */}
        <div className="lg:col-span-2 bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden">
          <div className="px-6 py-4 border-b border-slate-200">
            <h3 className="text-base font-semibold text-slate-900">Authorized Vendors</h3>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-sm">
              <thead className="bg-slate-100 text-slate-700 uppercase tracking-wider text-xs border-b border-slate-200">
                <tr>
                  <th className="py-3 px-6">Company</th>
                  <th className="py-3 px-6">Contact</th>
                  <th className="py-3 px-6">Email</th>
                  <th className="py-3 px-6 font-mono">System UID Mapping</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 text-slate-800">
                {suppliers.length === 0 ? (
                  <tr>
                    <td colSpan={4} className="py-8 text-center text-slate-400">
                      No suppliers mapped yet.
                    </td>
                  </tr>
                ) : (
                  suppliers.map((sup) => (
                    <tr key={sup.id} className="hover:bg-slate-50 transition-colors">
                      <td className="py-4 px-6 font-semibold text-slate-900">{sup.companyName}</td>
                      <td className="py-4 px-6">{sup.contactName || "—"}</td>
                      <td className="py-4 px-6 text-slate-600">{sup.email}</td>
                      <td className="py-4 px-6 font-mono text-xs text-slate-400 truncate max-w-[150px]">{sup.supplierId}</td>
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