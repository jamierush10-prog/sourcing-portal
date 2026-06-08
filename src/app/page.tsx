// src/app/page.tsx
import { redirect } from "next/navigation";

export default function RootPage() {
  // Automatically pass visitors straight to the authentication screen
  redirect("/login");
}