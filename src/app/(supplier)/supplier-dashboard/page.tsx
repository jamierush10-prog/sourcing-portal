// src/app/(supplier)/supplier-dashboard/page.tsx
"use client";

import React, { useState, useEffect, useMemo } from "react";
import ExcelJS from "exceljs";
import { collection, onSnapshot, query, where, doc, updateDoc, getDocs } from "firebase/firestore";
import { db, auth } from "@/lib/firebase"; 
import { useAuth } from "@/context/AuthContext";
import { useRouter } from "next/navigation";
import { signOut } from "firebase/auth";

interface RFQItem {
  id: string;
  materialId: string;
  rfqId: string; 
  itemNumber: string;
  description: string;
  quantity: number;
  uom: string;
  buyer: string; 
  status: "Pending" | "Completed";
  isHot?: boolean;
  offeredPrice: number | null;
  leadTime: string | null;
  supplierNote: string;
  timestamp: any; 
}

export default function SupplierDashboard() {
  const { profile, loading } = useAuth();
  const router = useRouter();

  const [rfqs, setRfqs] = useState<RFQItem[]>([]);
  const [materialsMap, setMaterialsMap] = useState<Record<string, any>>({});
  const [isDataLoading, setIsDataLoading] = useState(false);

  // Filter States
  const [isFilterModalOpen, setIsFilterModalOpen] = useState(false);
  const [filterRfqId, setFilterRfqId] = useState("");
  const [filterItemNumber, setFilterItemNumber] = useState("");
  const [filterDescription, setFilterDescription] = useState("");
  const [filterBuyer, setFilterBuyer] = useState("");

  const [editingId, setEditingId] = useState<string | null>(null);
  const [bidPrice, setBidPrice] = useState<string>("");
  const [leadTime, setLeadTime] = useState<string>("");
  const [vendorNotes, setVendorNotes] = useState<string>("");
  const [isSaving, setIsSaving] = useState(false);
  const [isGeneratingProposal, setIsGeneratingProposal] = useState(false);
  const [isExportingExcel, setIsExportingExcel] = useState(false);

  const filteredRows = useMemo(() => {
    return rfqs.filter((item) => {
      return (
        (item.rfqId || "").toLowerCase().includes(filterRfqId.trim().toLowerCase()) &&
        (item.itemNumber || "").toLowerCase().includes(filterItemNumber.trim().toLowerCase()) &&
        (item.description || "").toLowerCase().includes(filterDescription.trim().toLowerCase()) &&
        (item.buyer || "").toLowerCase().includes(filterBuyer.trim().toLowerCase())
      );
    });
  }, [rfqs, filterRfqId, filterItemNumber, filterDescription, filterBuyer]);

  useEffect(() => {
    if (!loading && (!profile || profile.role !== "supplier")) {
      router.push("/login");
    }
  }, [profile, loading, router]);

  useEffect(() => {
    const supplierProfile = profile as any;
    if (loading || profile?.role !== "supplier" || !supplierProfile?.supplierNo) return;

    setIsDataLoading(true);

    getDocs(collection(db, "materials")).then((materialsSnapshot) => {
      const matMap: Record<string, any> = {};
      materialsSnapshot.forEach((mDoc) => { matMap[mDoc.id] = mDoc.data(); });
      setMaterialsMap(matMap);
    }).catch(console.error);

    const routingQuery = query(collection(db, "rfq_routing"), where("supplierNo", "==", supplierProfile.supplierNo));
    const unsubscribe = onSnapshot(routingQuery, (snapshot) => {
      const list: RFQItem[] = [];
      snapshot.forEach((doc) => { list.push({ id: doc.id, ...doc.data() } as RFQItem); });
      setRfqs(list);
      setIsDataLoading(false);
    });
    return () => unsubscribe();
  }, [profile, loading]);

  const handleSupplierLogout = async () => {
    await signOut(auth);
    router.push("/login");
  };

  const handleSaveBid = async (rfqId: string) => {
    setIsSaving(true);
    try {
      await updateDoc(doc(db, "rfq_routing", rfqId), {
        offeredPrice: parseFloat(bidPrice) || 0,
        leadTime: leadTime.trim(),
        supplierNote: vendorNotes.trim(),
        status: "Completed",
        timestamp: new Date()
      });
      setEditingId(null);
    } finally { setIsSaving(false); }
  };

  const handleExportTableToExcel = async () => {
    setIsGeneratingProposal(true);
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Quote Proposal");
      
      // Header Section
      worksheet.mergeCells("A2:C2");
      worksheet.getCell("A2").value = (profile?.companyName || "VENDOR").toUpperCase();
      worksheet.getCell("A2").font = { size: 16, bold: true, color: { argb: "1E3A8A" } };
      worksheet.getCell("A6").value = `Supplier ID: ${(profile as any)?.supplierNo || "N/A"}`;
      worksheet.getCell("A7").value = `Contact: ${(profile as any)?.contactName || (profile as any)?.name || "N/A"}`;
      
      // Table
      const headers = ["RFQ ID", "Item #", "Description", "Qty", "Price", "Total"];
      worksheet.getRow(11).values = headers;
      filteredRows.forEach((item, i) => {
        const row = worksheet.addRow([item.rfqId, item.itemNumber, item.description, item.quantity, item.offeredPrice || 0, { formula: `=D${12+i}*E${12+i}` }]);
        row.eachCell((c) => c.alignment = { vertical: "center", horizontal: "center" });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const anchor = document.createElement("a");
      anchor.href = window.URL.createObjectURL(blob);
      anchor.download = `Quote_${profile?.companyName || "Proposal"}.xlsx`;
      anchor.click();
    } finally { setIsGeneratingProposal(false); }
  };

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <header className="mb-8 flex justify-between border-b pb-5">
        <div>
          <h1 className="text-3xl font-bold">{profile?.companyName || "Vendor"}</h1>
          <button onClick={handleSupplierLogout} className="text-red-600 text-xs font-bold mt-2">🚪 Logout</button>
        </div>
      </header>

      <div className="flex gap-2 mb-4">
        <button onClick={() => setIsFilterModalOpen(true)} className="bg-white border px-3 py-1 rounded text-sm">🔍 Filter</button>
        <button onClick={handleExportTableToExcel} className="bg-blue-600 text-white px-4 py-1 rounded text-sm font-bold">📑 Generate Quote Proposal</button>
      </div>

      <table className="w-full bg-white border rounded">
        <thead>
          <tr className="bg-slate-100 text-xs text-left">
            <th className="p-3">RFQ ID</th>
            <th className="p-3">Item #</th>
            <th className="p-3">Description</th>
            <th className="p-3">Qty</th>
            <th className="p-3">Price</th>
            <th className="p-3">Actions</th>
          </tr>
        </thead>
        <tbody className="divide-y">
          {filteredRows.map((item) => (
            <tr key={item.id} className={item.isHot ? "bg-red-50" : ""}>
              <td className="p-3">{item.isHot && "🔥 "}{item.rfqId}</td>
              <td className="p-3">{item.itemNumber}</td>
              <td className="p-3">{item.description}</td>
              <td className="p-3">{item.quantity}</td>
              <td className="p-3">${item.offeredPrice?.toFixed(2) || "0.00"}</td>
              <td className="p-3">
                <button onClick={() => setEditingId(item.id)} className="text-blue-600 font-bold text-xs">Edit</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}