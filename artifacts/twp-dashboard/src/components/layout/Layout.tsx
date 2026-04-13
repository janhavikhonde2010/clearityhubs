import { ReactNode } from "react";
import { Link, useLocation } from "wouter";
import { useCredentials } from "@/contexts/CredentialsContext";
import { CredentialsPrompt } from "./CredentialsPrompt";
import { 
  LayoutDashboard, 
  Users, 
  Headset, 
  MessageSquare,
  LogOut,
  Bird
} from "lucide-react";
import { Button } from "@/components/ui/button";

interface LayoutProps {
  children: ReactNode;
}

export function Layout({ children }: LayoutProps) {
  const { isConfigured, clearCredentials } = useCredentials();
  const [location] = useLocation();

  if (!isConfigured) {
    return <CredentialsPrompt />;
  }

  const navItems = [
    { href: "/", label: "Dashboard", icon: LayoutDashboard },
    { href: "/subscribers", label: "Subscribers", icon: Users },
    { href: "/agents", label: "Agents", icon: Headset },
    { href: "/sequences", label: "Sequences", icon: MessageSquare },
  ];

  return (
    <div className="flex min-h-[100dvh] bg-background text-foreground">
      {/* Sidebar */}
      <aside className="w-64 border-r border-border bg-card flex flex-col hidden md:flex shrink-0">
        <div className="h-16 flex items-center px-6 border-b border-border">
          <div className="flex items-center gap-2 text-primary">
            <Bird className="w-6 h-6" />
            <span className="font-bold tracking-tight text-lg">TWP Analytics</span>
          </div>
        </div>
        <nav className="flex-1 p-4 space-y-1">
          {navItems.map((item) => {
            const isActive = location === item.href;
            return (
              <Link key={item.href} href={item.href}>
                <div className={`flex items-center gap-3 px-3 py-2 rounded-md transition-colors cursor-pointer ${
                  isActive 
                    ? "bg-primary text-primary-foreground font-medium" 
                    : "text-muted-foreground hover:bg-secondary hover:text-secondary-foreground"
                }`}>
                  <item.icon className="w-4 h-4" />
                  {item.label}
                </div>
              </Link>
            );
          })}
        </nav>
        <div className="p-4 border-t border-border">
          <Button 
            variant="ghost" 
            className="w-full justify-start text-muted-foreground hover:text-destructive hover:bg-destructive/10"
            onClick={clearCredentials}
          >
            <LogOut className="w-4 h-4 mr-2" />
            Disconnect
          </Button>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 flex flex-col min-w-0 overflow-hidden">
        {/* Mobile Header */}
        <header className="h-16 flex items-center justify-between px-4 border-b border-border bg-card md:hidden shrink-0">
          <div className="flex items-center gap-2 text-primary">
            <Bird className="w-5 h-5" />
            <span className="font-bold tracking-tight">TWP</span>
          </div>
          <Button variant="ghost" size="icon" onClick={clearCredentials}>
            <LogOut className="w-4 h-4" />
          </Button>
        </header>

        {/* Scrollable Content Area */}
        <div className="flex-1 overflow-auto p-4 md:p-8">
          {children}
        </div>
      </main>
    </div>
  );
}
