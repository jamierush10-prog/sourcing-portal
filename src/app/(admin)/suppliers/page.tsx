// src/app/(admin)/suppliers/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import { collection, addDoc, getDocs, query, orderBy } from "firebase/firestore";
import { db } from "@/lib/firebase";
import { useAuth } from "@/context/AuthContext";
import { useRouter } from "next/navigation";

interface SupplierProfile {
  id: string;
  supplierId: string;
  companyName: string;
  contactName: string;
  email: string;
}

export default function SuppliersDirectory() {
  const { profile, loading } = useAuth();
  const router = useRouter();

  const [suppliers, setSuppliers] = useState<SupplierProfile[]>([]);
  const [companyName, setCompanyName] = useState("");
  const [contactName, setContactName] = useState("");
  const [email, setEmail] = useState("");
  const [supplierId, setSupplierId] = useState(""); // Firebase Authentication Auth UID mapping field
  
  const [isSubmitting, setIsSubmitting] = useState(false);

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

  const handleRegisterSupplier = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!companyName || !email || !supplierId) return;

    setIsSubmitting(true);
    try {
      await addDoc(collection(db, "suppliers"), {
        companyName,
        contactName,
        email,
        supplierId: supplierId.trim(),
      });

      setCompanyName("");
      setContactName("");
      setEmail("");
      setSupplierId("");
      fetchSuppliers();
    } catch (err) {
      console.error("Failed to append vendor record profiles:", err);
    } finally {
      setIsSubmitting(false);
    }
  };

  if (loading) return <div className="p-8">Verifying credentials...</div>;

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <header className="mb-8 flex justify-between items-center border-b border-slate-200 pb-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">Supplier Directory Configuration</h1>
          <p className="text-sm text-slate-500">Map active account profiles and execute isolated queue workspace simulations</p>
        </div>
        <button
          onClick={() => router.push("/dashboard")}
          className="text-sm font-semibold text-blue-600 hover:text-blue-800 bg-blue-50 px-3 py-1.5 rounded-md transition-colors"
        >
          ← Return to Console
        </button>
      </header>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
        {/* Vendor Profile Creation Card */}
        <div className="bg-white p-6 border border-slate-200 rounded-lg shadow-sm h-fit">
          <h3 className="text-sm font-bold uppercase text-slate-700 tracking-wider mb-4">Link Vendor Profile</h3>
          <form onSubmit={handleRegisterSupplier} className="space-y-4">
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
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">Firebase Account UID *</label>
              <input
                type="text"
                required
                value={supplierId}
                onChange={(e) => setSupplierId(e.target.value)}
                className="w-full text-xs font-mono rounded border border-slate-300 px-3 py-2 bg-slate-50 focus:outline-none focus:ring-1 focus:ring-blue-500"
                placeholder="Paste User UID from Firebase Auth"
              />
            </div>
            <button
              type="submit"
              disabled={isSubmitting}
              className="w-full rounded bg-blue-600 py-2 text-sm font-semibold text-white shadow-sm hover:bg-blue-500 transition-colors disabled:bg-blue-300"
            >
              {isSubmitting ? "Linking Profiles..." : "Register Vendor Link"}
            </button>
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
                  <th className="py-3 px-6">Company</th>
                  <th className="py-3 px-6">Contact</th>
                  <th className="py-3 px-6">Email</th>
                  <th className="py-3 px-6 text-center">System Action</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 text-slate-800">
                {suppliers.length === 0 ? (
                  <tr>
                    <td colSpan={4} className="py-8 text-center text-slate-400">
                      No suppliers registered in the database directory yet.
                    </td>
                  </tr>
                ) : (
                  suppliers.map((supplier) => (
                    <tr key={supplier.id} className="hover:bg-slate-50">
                      <td className="py-4 px-6 font-semibold text-slate-900">{supplier.companyName}</td>
                      <td className="py-4 px-6 text-slate-600">{supplier.contactName || "—"}</td>
                      <td className="py-4 px-6 text-slate-500 font-mono text-xs">{supplier.email}</td>
                      <td className="py-4 px-6 text-center">
                        <button
                          onClick={() => router.push(`/suppliers/${supplier.supplierId}`)}
                          className="inline-flex items-center rounded-md bg-white px-2.5 py-1.5 text-xs font-semibold text-slate-700 shadow-sm ring-1 ring-inset ring-slate-300 hover:bg-slate-50 transition-all"
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