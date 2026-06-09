"use client";

import React, { useState, useEffect, useMemo } from "react";
import ExcelJS from "exceljs";
import { collection, onSnapshot, query, where, doc, updateDoc } from "firebase/firestore";
import { db } from "@/lib/firebase";
import { useAuth } from "@/context/AuthContext";
import { useRouter } from "next/navigation";

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
  quoteDate?: any;
  quotedBy?: string;
}

export default function SupplierDashboard() {
  const { profile, loading } = useAuth();
  const router = useRouter();

  const [items, setItems] = useState<RFQItem[]>([]);
  const [isFilterModalOpen, setIsFilterModalOpen] = useState(false);
  const [isNoteModalOpen, setIsNoteModalOpen] = useState<{isOpen: boolean, itemId: string | null}>({isOpen: false, itemId: null});
  const [selectedItems, setSelectedItems] = useState<Set<string>>(new Set());
  
  const [filterRfqId, setFilterRfqId] = useState("");
  const [filterItem, setFilterItem] = useState("");
  const [filterDesc, setFilterDesc] = useState("");

  const [editingId, setEditingId] = useState<string | null>(null);
  const [bidPrice, setBidPrice] = useState("");
  const [bidLeadTime, setBidLeadTime] = useState("");
  const [bidNotes, setBidNotes] = useState("");

  const filteredItems = useMemo(() => {
    return items.filter(item => 
      item.rfqId.toLowerCase().includes(filterRfqId.toLowerCase()) &&
      item.itemNumber.toLowerCase().includes(filterItem.toLowerCase()) &&
      item.description.toLowerCase().includes(filterDesc.toLowerCase())
    );
  }, [items, filterRfqId, filterItem, filterDesc]);

  const clearFilters = () => { setFilterRfqId(""); setFilterItem(""); setFilterDesc(""); };
  const toggleSelect = (id: string) => { const next = new Set(selectedItems); next.has(id) ? next.delete(id) : next.add(id); setSelectedItems(next); };

  const handleSaveBid = async (id: string) => {
    await updateDoc(doc(db, "rfq_routing", id), {
      offeredPrice: parseFloat(bidPrice) || 0,
      leadTime: bidLeadTime,
      supplierNote: bidNotes,
      status: "Completed",
      quoteDate: new Date(),
      quotedBy: profile?.email || "Supplier"
    });
    setEditingId(null);
    setIsNoteModalOpen({isOpen: false, itemId: null});
  };

  return (
    <div className="min-h-screen bg-white p-8 text-black">
      <header className="border-b-4 border-black pb-4 mb-8">
        <h1 className="text-4xl font-black uppercase">Supplier Portal</h1>
      </header>

      <div className="flex gap-4 mb-6">
        <button onClick={() => setIsFilterModalOpen(true)} className="bg-black text-white px-6 py-2 font-black uppercase text-xs">Filter List</button>
      </div>

      <table className="w-full border-2 border-black text-left text-sm">
        <thead className="bg-black text-white uppercase font-black">
          <tr>
            <th className="p-3">RFQ ID</th>
            <th className="p-3">Description</th>
            <th className="p-3">Price</th>
            <th className="p-3">Actions</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-black">
          {filteredItems.map(item => (
            <tr key={item.id}>
              <td className="p-3 font-bold">{item.rfqId}</td>
              <td className="p-3">{item.description}</td>
              <td className="p-3">{editingId === item.id ? <input type="number" onChange={e => setBidPrice(e.target.value)} className="border p-1 w-20" /> : `$${item.offeredPrice || "0.00"}`}</td>
              <td className="p-3 flex gap-2">
                <button 
                  onClick={() => setIsNoteModalOpen({isOpen: true, itemId: item.id})}
                  onMouseEnter={() => setIsNoteModalOpen({isOpen: true, itemId: item.id})}
                  className={`px-3 py-1 font-bold border-2 border-black ${item.supplierNote ? 'bg-yellow-400' : 'bg-white'}`}>
                  NOTES
                </button>
                {editingId === item.id ? <button onClick={() => handleSaveBid(item.id)} className="font-bold underline">SAVE</button> : <button onClick={() => setEditingId(item.id)} className="font-bold underline">QUOTE ONLINE</button>}
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      {isNoteModalOpen.isOpen && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center">
          <div className="bg-white p-8 border-4 border-black w-80">
            <h2 className="font-black mb-4 uppercase">Item Notes</h2>
            <textarea className="w-full border-2 p-2 mb-4" rows={4} onChange={e => setBidNotes(e.target.value)} placeholder="Add note..."></textarea>
            <button onClick={() => handleSaveBid(isNoteModalOpen.itemId!)} className="w-full bg-black text-white py-2 font-bold">SAVE NOTE</button>
          </div>
        </div>
      )}
      {/* ... Filter Modal ... */}
    </div>
  );
}