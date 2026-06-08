// src/app/(admin)/users/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import { collection, doc, getDocs, setDoc, updateDoc, deleteDoc, query, orderBy } from "firebase/firestore";
import { db } from "@/lib/firebase";
import { useAuth } from "@/context/AuthContext";
import { useRouter } from "next/navigation";

interface UserAccount {
  id: string; // This corresponds to their Firebase Auth UID
  email: string;
  role: "admin" | "supplier";
  companyName?: string;
  supplierNo?: string;
}

interface SupplierRegistry {
  id: string;
  supplierNo: string;
  companyName: string;
}

export default function UserManagement() {
  const { profile, loading } = useAuth();
  const router = useRouter();

  const [users, setUsers] = useState<UserAccount[]>([]);
  const [suppliers, setSuppliers] = useState<SupplierRegistry[]>([]);
  const [isDataLoading, setIsDataLoading] = useState(true);

  // Form States (Create / Edit)
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [editingUser, setEditingUser] = useState<UserAccount | null>(null);
  const [formUid, setFormUid] = useState("");
  const [formEmail, setFormEmail] = useState("");
  const [formRole, setFormRole] = useState<"admin" | "supplier">("supplier");
  const [formCompanyName, setFormCompanyName] = useState("");
  
  // Modal State for Supplier Linking Assignment
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [activeModalUser, setActiveModalUser] = useState<UserAccount | null>(null);

  useEffect(() => {
    if (!loading && (!profile || profile.role !== "admin")) {
      router.push("/login");
    }
  }, [profile, loading, router]);

  useEffect(() => {
    if (profile?.role === "admin") {
      fetchUserRecords();
      fetchSupplierRegistry();
    }
  }, [profile]);

  const fetchUserRecords = async () => {
    setIsDataLoading(true);
    try {
      const snapshot = await getDocs(collection(db, "users"));
      const list: UserAccount[] = [];
      snapshot.forEach((doc) => {
        list.push({ id: doc.id, ...doc.data() } as UserAccount);
      });
      setUsers(list);
    } catch (err) {
      console.error("Error loading user collection records:", err);
    } finally {
      setIsDataLoading(false);
    }
  };

  const fetchSupplierRegistry = async () => {
    try {
      const q = query(collection(db, "suppliers"), orderBy("companyName", "asc"));
      const snapshot = await getDocs(q);
      const list: SupplierRegistry[] = [];
      snapshot.forEach((doc) => {
        list.push({ id: doc.id, ...doc.data() } as SupplierRegistry);
      });
      setSuppliers(list);
    } catch (err) {
      console.error("Error loading profile registries:", err);
    }
  };

  const handleOpenCreateForm = () => {
    setEditingUser(null);
    setFormUid("");
    setFormEmail("");
    setFormRole("supplier");
    setFormCompanyName("");
    setIsFormOpen(true);
  };

  const handleOpenEditForm = (user: UserAccount) => {
    setEditingUser(user);
    setFormUid(user.id);
    setFormEmail(user.email || "");
    setFormRole(user.role);
    setFormCompanyName(user.companyName || "");
    setIsFormOpen(true);
  };

  const handleSaveUserDocument = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formUid.trim() || !formEmail.trim()) return;

    try {
      const userDocRef = doc(db, "users", formUid.trim());
      const payload = {
        email: formEmail.trim(),
        role: formRole,
        companyName: formRole === "admin" ? "Internal Admin" : formCompanyName.trim(),
      };

      if (editingUser) {
        await updateDoc(userDocRef, payload);
      } else {
        // Create new record mapping profile parameters
        await setDoc(userDocRef, {
          ...payload,
          supplierNo: "" // Initially unlinked
        });
      }

      setIsFormOpen(false);
      fetchUserRecords();
    } catch (err) {
      console.error("Failed to commit user configuration mapping:", err);
    }
  };

  const handleDeleteUserDocument = async (uid: string) => {
    if (!confirm("Are you sure you want to completely remove this user profile entry from Firestore?")) return;
    try {
      await deleteDoc(doc(db, "users", uid));
      fetchUserRecords();
    } catch (err) {
      console.error("Failed to purge target profile record:", err);
    }
  };

  const handleOpenLinkModal = (user: UserAccount) => {
    setActiveModalUser(user);
    setIsModalOpen(true);
  };

  const handleAssignSupplierMapping = async (supplier: SupplierRegistry) => {
    if (!activeModalUser) return;

    try {
      const userDocRef = doc(db, "users", activeModalUser.id);
      
      // Update the user record with the company metadata parameters
      await updateDoc(userDocRef, {
        supplierNo: supplier.supplierNo,
        companyName: supplier.companyName
      });

      setIsModalOpen(false);
      setActiveModalUser(null);
      fetchUserRecords(); // Reload grid metrics
    } catch (err) {
      console.error("Failed to map target vendor permissions key:", err);
    }
  };

  if (loading) return <div className="p-8 text-sm text-slate-500">Verifying security parameters...</div>;

  return (
    <div className="min-h-screen p-8 bg-slate-50">
      <header className="mb-8 flex justify-between items-center border-b border-slate-200 pb-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">User Accounts Management</h1>
          <p className="text-sm text-slate-500">Provision authorization profiles and map secure vendor supplier access keys</p>
        </div>
        <div className="flex gap-3">
          <button
            onClick={() => router.push("/dashboard")}
            className="text-sm font-semibold text-slate-600 bg-white border border-slate-300 px-3 py-1.5 rounded-md hover:bg-slate-50"
          >
            ← Main Console
          </button>
          <button
            onClick={handleOpenCreateForm}
            className="text-sm font-semibold text-white bg-blue-600 px-3 py-1.5 rounded-md hover:bg-blue-500 shadow-sm"
          >
            + Add New User Profile
          </button>
        </div>
      </header>

      {/* Main Grid View */}
      <div className="bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden">
        <table className="w-full text-left border-collapse text-sm">
          <thead className="bg-slate-100 text-slate-700 font-semibold text-xs border-b border-slate-200">
            <tr>
              <th className="py-3 px-6">Account UID</th>
              <th className="py-3 px-6">Email Profile</th>
              <th className="py-3 px-6">Role Privileges</th>
              <th className="py-3 px-6">Mapped Company</th>
              <th className="py-3 px-6">Linked Supplier No</th>
              <th className="py-3 px-6 text-center">System Management Actions</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200 text-slate-800">
            {isDataLoading ? (
              <tr><td colSpan={6} className="py-8 text-center text-slate-400">Loading master account profiles...</td></tr>
            ) : users.length === 0 ? (
              <tr><td colSpan={6} className="py-8 text-center text-slate-400">No user tracking entries defined yet.</td></tr>
            ) : (
              users.map((u) => (
                <tr key={u.id} className="hover:bg-slate-50">
                  <td className="py-4 px-6 font-mono text-xs text-slate-500 truncate max-w-[120px]" title={u.id}>{u.id}</td>
                  <td className="py-4 px-6 font-medium text-slate-900">{u.email}</td>
                  <td className="py-4 px-6">
                    <span className={`inline-flex items-center rounded-md px-2 py-0.5 text-xs font-medium ${u.role === 'admin' ? 'bg-purple-50 text-purple-700 ring-1 ring-purple-700/10' : 'bg-blue-50 text-blue-700 ring-1 ring-blue-700/10'}`}>
                      {u.role.toUpperCase()}
                    </span>
                  </td>
                  <td className="py-4 px-6 text-slate-700">{u.companyName || "—"}</td>
                  <td className="py-4 px-6">
                    {u.supplierNo ? (
                      <span className="font-mono font-bold text-blue-600 bg-blue-50 px-2 py-0.5 rounded">{u.supplierNo}</span>
                    ) : u.role === "supplier" ? (
                      <span className="text-amber-600 font-semibold text-xs bg-amber-50 px-2 py-0.5 rounded animate-pulse">Unlinked Access</span>
                    ) : (
                      <span className="text-slate-300">—</span>
                    )}
                  </td>
                  <td className="py-4 px-6 text-center whitespace-nowrap space-x-2">
                    {u.role === "supplier" && (
                      <button
                        onClick={() => handleOpenLinkModal(u)}
                        className="bg-emerald-50 text-emerald-700 hover:bg-emerald-100 text-xs font-semibold px-2.5 py-1.5 rounded transition-colors"
                      >
                        🔗 Link Supplier No
                      </button>
                    )}
                    <button
                      onClick={() => handleOpenEditForm(u)}
                      className="border border-slate-300 text-slate-700 hover:bg-slate-50 text-xs font-semibold px-2.5 py-1.5 rounded transition-all"
                    >
                      Edit
                    </button>
                    <button
                      onClick={() => handleDeleteUserDocument(u.id)}
                      className="bg-red-50 text-red-600 hover:bg-red-100 text-xs font-semibold px-2.5 py-1.5 rounded transition-colors"
                    >
                      Delete
                    </button>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {/* CREATE / EDIT ACCOUNT SLIDEOUT MODAL */}
      {isFormOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="w-full max-w-md rounded-lg bg-white p-6 shadow-xl border border-slate-200">
            <h3 className="text-lg font-bold text-slate-900 mb-4">{editingUser ? "Edit Profile Settings" : "Provision New User Profile Document"}</h3>
            <form onSubmit={handleSaveUserDocument} className="space-y-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Firebase Auth User UID *</label>
                <input
                  type="text"
                  required
                  disabled={!!editingUser}
                  value={formUid}
                  onChange={(e) => setFormUid(e.target.value)}
                  className="w-full text-xs font-mono rounded border border-slate-300 px-3 py-2 focus:outline-none focus:ring-1 focus:ring-blue-500 bg-slate-50 disabled:text-slate-400"
                  placeholder="Paste long string from Firebase Auth screen"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Email Address *</label>
                <input
                  type="email"
                  required
                  value={formEmail}
                  onChange={(e) => setFormEmail(e.target.value)}
                  className="w-full text-sm rounded border border-slate-300 px-3 py-2 focus:outline-none focus:ring-1 focus:ring-blue-500"
                  placeholder="rep@supplier.com"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">System Authorization Role *</label>
                <select
                  value={formRole}
                  onChange={(e) => setFormRole(e.target.value as any)}
                  className="w-full text-sm rounded border border-slate-300 px-3 py-2 focus:outline-none focus:ring-1 focus:ring-blue-500 cursor-pointer"
                >
                  <option value="supplier">Supplier Portal Access</option>
                  <option value="admin">System Administrator Console</option>
                </select>
              </div>
              {formRole === "supplier" && !editingUser && (
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Temporary Label Company Name</label>
                  <input
                    type="text"
                    value={formCompanyName}
                    onChange={(e) => setFormCompanyName(e.target.value)}
                    className="w-full text-sm rounded border border-slate-300 px-3 py-2 focus:outline-none focus:ring-1 focus:ring-blue-500"
                    placeholder="e.g. Temp Company Name"
                  />
                </div>
              )}
              <div className="flex justify-end gap-3 border-t border-slate-200 pt-4 mt-6">
                <button type="button" onClick={() => setIsFormOpen(false)} className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50">Cancel</button>
                <button type="submit" className="rounded-md bg-blue-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-blue-500/90">Save Configuration</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* SCROLLABLE SUPPLIER LINK SELECTION MODAL */}
      {isModalOpen && activeModalUser && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="w-full max-w-md rounded-lg bg-white p-6 shadow-xl border border-slate-200 flex flex-col max-h-[80vh]">
            <div className="border-b border-slate-200 pb-3 mb-4">
              <h3 className="text-lg font-bold text-slate-900">Map Vendor Data Clearance</h3>
              <p className="text-xs text-slate-500 mt-1">Assign matching data visibility rules for account: <span className="font-semibold text-slate-700">{activeModalUser.email}</span></p>
            </div>
            
            {/* Scrollable List Container */}
            <div className="overflow-y-auto flex-1 space-y-2 pr-1 my-2 min-h-[200px]">
              {suppliers.length === 0 ? (
                <div className="p-8 text-center text-xs text-slate-400">No active companies found in the Suppliers list. Register companies inside /suppliers directory page first.</div>
              ) : (
                suppliers.map((sup) => (
                  <div
                    key={sup.id}
                    onClick={() => handleAssignSupplierMapping(sup)}
                    className="flex items-center justify-between p-3 rounded-md border border-slate-200 hover:border-blue-500 hover:bg-blue-50/40 cursor-pointer transition-all"
                  >
                    <div>
                      <p className="text-sm font-semibold text-slate-900">{sup.companyName}</p>
                    </div>
                    <span className="font-mono font-bold text-xs text-blue-700 bg-blue-50 px-2 py-1 rounded">
                      {sup.supplierNo}
                    </span>
                  </div>
                ))
              )}
            </div>

            <div className="flex justify-end border-t border-slate-200 pt-4 mt-4">
              <button
                type="button"
                onClick={() => { setIsModalOpen(false); setActiveModalUser(null); }}
                className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50"
              >
                Close Window
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}