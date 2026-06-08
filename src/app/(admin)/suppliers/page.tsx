// src/app/(admin)/suppliers/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import { collection, addDoc, getDocs, query, orderBy } from "firebase/firestore";
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

export default function SuppliersManagement() {
  const { profile, loading } = useAuth();
  const router = useRouter();

  const [suppliers, setSuppliers] = useState<SupplierProfile[]>([]);
  const [isDataLoading, setIsDataLoading] = useState(false);

  // Form Input States
  const [supplierNo, setSupplierNo] = useState("");
  const [companyName, setCompanyName] = useState("");
  const [contactName, setContactName] = useState("");
  const [email, setEmail] = useState("");
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [formFeedback, setFormFeedback] = useState("");

  useEffect(() => {
    if (!loading && (!profile || profile.role !== "admin")) {
      router.push("/login");
    }
  }, [profile, loading, router]);

  useEffect(() => {
    if (profile?.role === "admin") {
      fetchSuppliersList();
    }
  }, [profile]);

  const fetchSuppliersList = async () => {
    setIsDataLoading(true);
    try {
      const q = query(collection(db, "suppliers"), orderBy("companyName", "asc"));
      const snapshot = await getDocs(q);
      const list: SupplierProfile[] = [];
      snapshot.forEach((doc) => {
        list.push({ id: doc.id, ...doc.data() } as SupplierProfile);
      });
      setSuppliers(list);
    } catch (err) {
      console.error("Error pulling corporate vendor files:", err);
    } finally {
      setIsDataLoading(false);
    }
  };

  const handleRegisterVendorLink = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!supplierNo.trim() || !companyName.trim() || !email.trim()) {
      setFormFeedback("Please fill out all required fields (*).");
      return;
    }

    if (supplierNo.trim().length < 4 || supplierNo.trim().length > 6) {
      setFormFeedback("Supplier number must be between 4 and 6 characters.");
      return;
    }

    setIsSubmitting(true);
    setFormFeedback("");

    try {
      await addDoc(collection(db, "suppliers"), {
        supplierNo: supplierNo.trim().toUpperCase(),
        companyName: companyName.trim(),
        contactName: contactName.trim(),
        email: email.trim().toLowerCase(),
        timestamp: new Date()
      });

      setSupplierNo("");
      setCompanyName("");
      setContactName("");
      setEmail("");
      setFormFeedback("Vendor profile linked successfully!");
      fetchSuppliersList();
    } catch (err) {
      console.error("Failed storing vendor profile reference:", err);
      setFormFeedback("Database write failure. Verify connection parameters.");
    } finally {
      setIsSubmitting(false);
    }
  };

  if (loading) return <div className="p-8 text-sm text-slate-500">Verifying administrative access credentials...</div>;

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <header className="mb-8 flex justify-between items-center border-b border-slate-200 pb-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">Corporate Suppliers Directory</h1>
          <p className="text-sm text-slate-500">Manage vendor parameter mappings and link unique profiles</p>
        </div>
        <button
          onClick={() => router.push("/dashboard")}
          className="text-sm font-semibold text-slate-700 bg-white border border-slate-300 px-4 py-2 rounded-md hover:bg-slate-50 shadow-sm transition-all"
        >
          ⬅️ Back to Console
        </button>
      </header>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8 items-start">
        {/* LINK VENDOR PROFILE FORM CONTAINER */}
        <div className="bg-white p-6 rounded-lg border border-slate-200 shadow-sm">
          <h3 className="text-lg font-bold uppercase tracking-wide text-slate-800 mb-6 border-b border-slate-100 pb-2">
            Link Vendor Profile
          </h3>
          
          <form onSubmit={handleRegisterVendorLink} className="space-y-4">
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">Supplier No * (4-6 Chars)</label>
              <input
                type="text"
                placeholder="E.G. PSTN"
                value={supplierNo}
                onChange={(e) => setSupplierNo(e.target.value)}
                className="w-full rounded border border-slate-300 px-3 py-2 text-sm text-slate-900 placeholder:text-slate-300 focus:outline-none focus:ring-1 focus:ring-blue-500 font-mono font-bold uppercase"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">Company Name *</label>
              <input
                type="text"
                placeholder="Piston Supply Co."
                value={companyName}
                onChange={(e) => setCompanyName(e.target.value)}
                className="w-full rounded border border-slate-300 px-3 py-2 text-sm text-slate-900 placeholder:text-slate-300 focus:outline-none focus:ring-1 focus:ring-blue-500 font-medium"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">Primary Contact Name</label>
              <input
                type="text"
                placeholder="John Doe"
                value={contactName}
                onChange={(e) => setContactName(e.target.value)}
                className="w-full rounded border border-slate-300 px-3 py-2 text-sm text-slate-900 placeholder:text-slate-300 focus:outline-none focus:ring-1 focus:ring-blue-500"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">Email Address *</label>
              <input
                type="email"
                placeholder="sales@piston.com"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                className="w-full rounded border border-slate-300 px-3 py-2 text-sm text-slate-900 placeholder:text-slate-300 focus:outline-none focus:ring-1 focus:ring-blue-500"
              />
            </div>

            {formFeedback && (
              <p className={`text-xs font-semibold p-2 rounded ${formFeedback.includes("successfully") ? 'bg-emerald-50 text-emerald-800' : 'bg-red-50 text-red-800'}`}>
                {formFeedback}
              </p>
            )}

            <button
              type="submit"
              disabled={isSubmitting}
              className="w-full rounded-md bg-blue-600 py-2.5 text-sm font-semibold text-white shadow-sm hover:bg-blue-500 disabled:bg-blue-300 transition-colors mt-2"
            >
              {isSubmitting ? "Linking Record..." : "Register Vendor Link"}
            </button>
          </form>
        </div>

        {/* ACTIVE DIRECTORY TABLE QUEUE */}
        <div className="lg:col-span-2 bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden">
          <div className="px-6 py-4 border-b border-slate-200 bg-slate-50/70">
            <h3 className="text-sm font-bold uppercase text-slate-700 tracking-wider">Active Supplier Registry</h3>
          </div>
          <div className="overflow-x-auto">
            {isDataLoading ? (
              <div className="p-12 text-center text-slate-400">Querying directory databases...</div>
            ) : suppliers.length === 0 ? (
              <div className="p-12 text-center text-slate-400">No vendor profiles mapped yet.</div>
            ) : (
              <table className="w-full text-left border-collapse text-sm">
                <thead className="bg-slate-100 text-slate-700 font-semibold text-xs border-b border-slate-200">
                  <tr>
                    <th className="py-3 px-6 text-center">Code</th>
                    <th className="py-3 px-6">Company Name</th>
                    <th className="py-3 px-6">Primary Contact</th>
                    <th className="py-3 px-6">Email Destination</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 text-slate-800">
                  {suppliers.map((supplier) => (
                    <tr key={supplier.id} className="hover:bg-slate-50/60 transition-colors">
                      <td className="py-4 px-6 text-center font-mono font-bold text-xs text-blue-600 bg-blue-50/30">{supplier.supplierNo}</td>
                      <td className="py-4 px-6 font-semibold text-slate-900">{supplier.companyName}</td>
                      <td className="py-4 px-6 font-medium text-slate-700">{supplier.contactName || <span className="text-slate-300">—</span>}</td>
                      <td className="py-4 px-6 text-slate-500 font-mono text-xs">{supplier.email}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}