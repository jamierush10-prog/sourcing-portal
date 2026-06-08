// src/app/(auth)/login/page.tsx
"use client";

import React, { useState, useEffect } from "react";
import { signInWithEmailAndPassword } from "firebase/auth";
import { useRouter } from "next/navigation";
import { auth } from "@/lib/firebase";
import { useAuth } from "@/context/AuthContext";

export default function LoginPage() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [isSubmitting, setIsSubmitting] = useState(false);
  
  const router = useRouter();
  const { profile, loading } = useAuth();

  // Redirect users automatically if they are already logged in
  useEffect(() => {
    if (!loading && profile) {
      if (profile.role === "admin") {
        router.push("/dashboard");
      } else if (profile.role === "supplier") {
        router.push("/supplier-dashboard");
      }
    }
  }, [profile, loading, router]);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setError("");
    setIsSubmitting(true);

    try {
      await signInWithEmailAndPassword(auth, email, password);
      // The useEffect above will handle redirection once the auth state updates
    } catch (err: any) {
      console.error(err);
      setError(err.message || "Invalid email or password. Please try again.");
      setIsSubmitting(false);
    }
  };

  if (loading) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-slate-50">
        <p className="text-slate-600">Loading system parameters...</p>
      </div>
    );
  }

  return (
    <div className="flex min-h-screen items-center justify-center bg-slate-100 px-4 py-12 sm:px-6 lg:px-8">
      <div className="w-full max-w-md space-y-8 rounded-xl bg-white p-8 shadow-md border border-slate-200">
        <div>
          <h2 className="mt-2 text-center text-3xl font-bold tracking-tight text-slate-900">
            Sourcing Portal
          </h2>
          <p className="mt-2 text-center text-sm text-slate-600">
            Sign in to manage material sourcing and quotes
          </p>
        </div>
        
        <form className="mt-8 space-y-6" onSubmit={handleLogin}>
          {error && (
            <div className="rounded-md bg-red-50 p-3 border border-red-200">
              <p className="text-sm text-red-700">{error}</p>
            </div>
          )}
          
          <div className="space-y-4 rounded-md shadow-sm">
            <div>
              <label htmlFor="email-address" className="block text-sm font-medium text-slate-700 mb-1">
                Email Address
              </label>
              <input
                id="email-address"
                name="email"
                type="email"
                autoComplete="email"
                required
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                className="relative block w-full rounded-md border border-slate-300 px-3 py-2 text-slate-900 placeholder-slate-400 focus:z-10 focus:border-blue-500 focus:outline-none focus:ring-blue-500 sm:text-sm"
                placeholder="chief@company.com"
              />
            </div>
            
            <div>
              <label htmlFor="password" className="block text-sm font-medium text-slate-700 mb-1">
                Password
              </label>
              <input
                id="password"
                name="password"
                type="password"
                autoComplete="current-password"
                required
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                className="relative block w-full rounded-md border border-slate-300 px-3 py-2 text-slate-900 placeholder-slate-400 focus:z-10 focus:border-blue-500 focus:outline-none focus:ring-blue-500 sm:text-sm"
                placeholder="••••••••"
              />
            </div>
          </div>

          <div>
            <button
              type="submit"
              disabled={isSubmitting}
              className="group relative flex w-full justify-center rounded-md bg-blue-600 px-3 py-2 text-sm font-semibold text-white hover:bg-blue-700 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-blue-600 disabled:bg-blue-400 transition-colors"
            >
              {isSubmitting ? "Authenticating..." : "Sign In"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}