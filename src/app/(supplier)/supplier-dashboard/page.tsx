"use client";

import React, { useState, useEffect, useMemo } from "react";
import ExcelJS from "exceljs";
import { collection, onSnapshot, query, where, doc, updateDoc } from "firebase/firestore";
import { db } from "@/lib/firebase";
import { useAuth } from "@/context/AuthContext";
import { useRouter } from "next/navigation";
import { signOut } from "firebase/auth";

interface RFQItem {
  id: string;
  rfqId: string;
  itemNumber: string;
  description: string;
  quantity: number;
  uom: string;
  buyer: string;
  offeredPrice: number | null;
  leadTime: string | null;
  supplierNote: string;
  status: "Pending" | "Completed";
  supplierNo: string;
}

export default function SupplierDashboard() {
  const { profile, loading } = useAuth();
  const router = useRouter();

  const [items, setItems] = useState<RFQItem[]>([]);
  const [isFilterModalOpen, setIsFilterModalOpen] = useState(false);
  const [selectedItems, setSelectedItems] = useState<Set<string>>(new Set());
  
  // Filter states
  const [filterRfqId, setFilterRfqId] = useState("");
  const [filterItem, setFilterItem] = useState("");
  const [filterDesc, setFilterDesc] = useState("");

  // Edit states
  const [editingId, setEditingId] = useState<string | null>(null);
  const [bidPrice, setBidPrice] = useState("");

  const filteredItems = useMemo(() => {
    return items.filter(item => 
      item.rfqId.toLowerCase().includes(filterRfqId.toLowerCase()) &&
      item.itemNumber.toLowerCase().includes(filterItem.toLowerCase()) &&
      item.description.toLowerCase().includes(filterDesc.toLowerCase())
    );
  }, [items, filterRfqId, filterItem, filterDesc]);

  useEffect(() => {
    if (!loading && (!profile || profile.role !== "supplier")) router.push("/login");
  }, [profile, loading, router]);

  useEffect(() => {
    const p = profile as any;
    if (!p?.supplierNo) return;
    const q = query(collection(db, "rfq_routing"), where("supplierNo", "==", p.supplierNo));
    return onSnapshot(q, (snap) => {
      const list: RFQItem[] = [];
      snap.forEach(d => list.push({ id: d.id, ...d.data() } as RFQItem));
      setItems(list);
    });
  }, [profile]);

  const toggleSelect = (id: string) => {
    const next = new Set(selectedItems);
    if (next.has(id)) next.delete(id); else next.add(id);
    setSelectedItems(next);
  };

  const handleInlineSave = async (id: string) => {
    await updateDoc(doc(db, "rfq_routing", id), {
      offeredPrice: parseFloat(bidPrice),
      status: "Completed"
    });
    setEditingId(null);
  };

  const exportSelectedToExcel = async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Quote");
    ws.columns = [
      { header: "RFQ ID", key: "rfqId", width: 15 },
      { header: "Item #", key: "item", width: 15 },
      { header: "Description", key: "desc", width: 40 },
      { header: "Qty", key: "qty", width: 10 },
      { header: "Price", key: "price", width: 15 }
    ];
    
    filteredItems.filter(i => selectedItems.has(i.id)).forEach(i => {
      ws.addRow({ rfqId: i.rfqId, item: i.itemNumber, desc: i.description, qty: i.quantity, price: i.offeredPrice });
    });
    
    const buf = await wb.xlsx.writeBuffer();
    const url = window.URL.createObjectURL(new Blob([buf]));
    const a = document.createElement("a");
    a.href = url;
    a.download = "Quote.xlsx";
    a.click();
  };

  return (
    <div className="min-h-screen bg-white p-8 text-black">
      <header className="border-b-4 border-black pb-4 mb-8">
        <h1 className="text-4xl font-black uppercase">Supplier Portal</h1>
        <p className="text-sm font-bold mt-1">{profile?.companyName}</p>
      </header>

      <div className="flex gap-4 mb-6">
        <button onClick={() => setIsFilterModalOpen(true)} className="bg-black text-white px-6 py-2 font-black uppercase text-xs">Filter List</button>
        <button onClick={exportSelectedToExcel} className="border-2 border-black px-6 py-2 font-black uppercase text-xs">Export Selected</button>
      </div>

      <table className="w-full border-2 border-black text-left">
        <thead className="bg-black text-white uppercase text-xs font-black">
          <tr>
            <th className="p-3">Select</th>
            <th className="p-3">RFQ ID</th>
            <th className="p-3">Item #</th>
            <th className="p-3">Description</th>
            <th className="p-3">Price</th>
            <th className="p-3">Action</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-black">
          {filteredItems.map(item => (
            <tr key={item.id}>
              <td className="p-3"><input type="checkbox" onChange={() => toggleSelect(item.id)} /></td>
              <td className="p-3 font-bold">{item.rfqId}</td>
              <td className="p-3 font-bold">{item.itemNumber}</td>
              <td className="p-3">{item.description}</td>
              <td className="p-3">
                {editingId === item.id ? (
                  <input type="number" onChange={(e) => setBidPrice(e.target.value)} className="border p-1" />
                ) : `$${item.offeredPrice || "0.00"}`}
              </td>
              <td className="p-3">
                {editingId === item.id ? (
                  <button onClick={() => handleInlineSave(item.id)} className="font-bold underline">SAVE</button>
                ) : (
                  <button onClick={() => setEditingId(item.id)} className="font-bold underline">EDIT</button>
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      {isFilterModalOpen && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center">
          <div className="bg-white p-8 border-4 border-black w-96">
            <h2 className="font-black mb-4">FILTER ITEMS</h2>
            <input placeholder="RFQ ID" className="w-full border-2 p-2 mb-2" onChange={e => setFilterRfqId(e.target.value)} />
            <input placeholder="Item #" className="w-full border-2 p-2 mb-2" onChange={e => setFilterItem(e.target.value)} />
            <input placeholder="Description" className="w-full border-2 p-2 mb-4" onChange={e => setFilterDesc(e.target.value)} />
            <button onClick={() => setIsFilterModalOpen(false)} className="w-full bg-black text-white py-2 font-bold">APPLY</button>
          </div>
        </div>
      )}
    </div>
  );
}