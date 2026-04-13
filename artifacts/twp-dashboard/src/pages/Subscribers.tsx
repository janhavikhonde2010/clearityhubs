import { useState, useMemo } from "react";
import { useGetSubscribers } from "@workspace/api-client-react";
import { useCredentials } from "@/contexts/CredentialsContext";
import { Input } from "@/components/ui/input";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import { Skeleton } from "@/components/ui/skeleton";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Search, Filter } from "lucide-react";
import { format } from "date-fns";

export default function Subscribers() {
  const { credentials } = useCredentials();
  
  const queryOpts = {
    apiToken: credentials?.apiToken || "",
    phoneNumberId: credentials?.phoneNumberId || ""
  };

  const { data, isLoading } = useGetSubscribers(queryOpts, {
    query: { enabled: !!credentials?.apiToken }
  });

  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState<string>("all");
  const [agentFilter, setAgentFilter] = useState<string>("all");
  const [labelFilter, setLabelFilter] = useState<string>("all");

  const subscribers = data?.subscribers || [];

  const agents = useMemo(() => {
    const unique = new Set(subscribers.map(s => s.assignedAgent).filter(Boolean));
    return Array.from(unique) as string[];
  }, [subscribers]);

  const labels = useMemo(() => {
    const unique = new Set(subscribers.map(s => s.labelName).filter(Boolean));
    return Array.from(unique) as string[];
  }, [subscribers]);

  const filteredSubscribers = useMemo(() => {
    return subscribers.filter(sub => {
      const matchesSearch = 
        sub.name.toLowerCase().includes(search.toLowerCase()) || 
        sub.phoneNumber.includes(search);
        
      const matchesStatus = 
        statusFilter === "all" ? true :
        statusFilter === "dormant" ? sub.isDormant :
        !sub.isDormant;

      const matchesAgent = agentFilter === "all" || sub.assignedAgent === agentFilter;
      const matchesLabel = labelFilter === "all" || sub.labelName === labelFilter;

      return matchesSearch && matchesStatus && matchesAgent && matchesLabel;
    });
  }, [subscribers, search, statusFilter, agentFilter, labelFilter]);

  return (
    <div className="space-y-6 max-w-full">
      <div>
        <h1 className="text-3xl font-bold tracking-tight">Subscribers Explorer</h1>
        <p className="text-muted-foreground mt-1">Detailed view of all leads and their conversation status.</p>
      </div>

      <div className="bg-card border border-border p-4 rounded-lg flex flex-col md:flex-row gap-4 items-start md:items-center justify-between shadow-sm">
        <div className="relative flex-1 max-w-md">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground" />
          <Input 
            placeholder="Search by name or phone..." 
            className="pl-9 bg-background"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
        </div>
        
        <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
          <div className="flex items-center gap-2 text-sm font-medium text-muted-foreground mr-1">
            <Filter className="h-4 w-4" /> Filters:
          </div>
          
          <Select value={statusFilter} onValueChange={setStatusFilter}>
            <SelectTrigger className="w-[130px] bg-background">
              <SelectValue placeholder="Status" />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="all">All Statuses</SelectItem>
              <SelectItem value="active">Active</SelectItem>
              <SelectItem value="dormant">Dormant</SelectItem>
            </SelectContent>
          </Select>

          <Select value={agentFilter} onValueChange={setAgentFilter}>
            <SelectTrigger className="w-[150px] bg-background">
              <SelectValue placeholder="Agent" />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="all">All Agents</SelectItem>
              {agents.map(a => (
                <SelectItem key={a} value={a}>{a}</SelectItem>
              ))}
            </SelectContent>
          </Select>

          <Select value={labelFilter} onValueChange={setLabelFilter}>
            <SelectTrigger className="w-[150px] bg-background">
              <SelectValue placeholder="Label" />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="all">All Labels</SelectItem>
              {labels.map(l => (
                <SelectItem key={l} value={l}>{l}</SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>
      </div>

      <div className="border border-border rounded-lg overflow-hidden bg-card shadow-sm">
        <div className="overflow-x-auto">
          <Table>
            <TableHeader className="bg-muted/50">
              <TableRow>
                <TableHead className="w-[200px]">Subscriber</TableHead>
                <TableHead>Status</TableHead>
                <TableHead>Label</TableHead>
                <TableHead>Agent</TableHead>
                <TableHead>Sequence</TableHead>
                <TableHead className="text-right">Replies (U/T)</TableHead>
                <TableHead className="text-right">Last Activity</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {isLoading ? (
                Array.from({ length: 10 }).map((_, i) => (
                  <TableRow key={i}>
                    <TableCell><Skeleton className="h-5 w-32" /><Skeleton className="h-3 w-24 mt-1" /></TableCell>
                    <TableCell><Skeleton className="h-5 w-16" /></TableCell>
                    <TableCell><Skeleton className="h-5 w-20" /></TableCell>
                    <TableCell><Skeleton className="h-5 w-24" /></TableCell>
                    <TableCell><Skeleton className="h-5 w-32" /></TableCell>
                    <TableCell className="text-right"><Skeleton className="h-5 w-12 ml-auto" /></TableCell>
                    <TableCell className="text-right"><Skeleton className="h-5 w-24 ml-auto" /></TableCell>
                  </TableRow>
                ))
              ) : filteredSubscribers.length === 0 ? (
                <TableRow>
                  <TableCell colSpan={7} className="h-32 text-center text-muted-foreground">
                    No subscribers found matching the filters.
                  </TableCell>
                </TableRow>
              ) : (
                filteredSubscribers.map((sub, idx) => (
                  <TableRow key={`${sub.phoneNumber}-${idx}`} className="hover:bg-muted/30">
                    <TableCell>
                      <div className="font-medium text-foreground">{sub.name}</div>
                      <div className="text-xs text-muted-foreground font-mono">{sub.phoneNumber}</div>
                    </TableCell>
                    <TableCell>
                      {sub.isDormant ? (
                        <Badge variant="destructive" className="bg-destructive/10 text-destructive border-destructive/20 hover:bg-destructive/20">Dormant</Badge>
                      ) : (
                        <Badge variant="outline" className="bg-primary/10 text-primary border-primary/20 hover:bg-primary/20">Active</Badge>
                      )}
                    </TableCell>
                    <TableCell>
                      <Badge variant="secondary" className="bg-secondary/50 font-normal">{sub.labelName}</Badge>
                    </TableCell>
                    <TableCell className="text-sm">
                      {sub.assignedAgent || <span className="text-muted-foreground italic">Unassigned</span>}
                    </TableCell>
                    <TableCell className="text-sm truncate max-w-[150px]" title={sub.assignedSequence || undefined}>
                      {sub.assignedSequence || <span className="text-muted-foreground italic">None</span>}
                    </TableCell>
                    <TableCell className="text-right font-mono text-sm">
                      <span className="text-primary">{sub.userReplyCount}</span>
                      <span className="text-muted-foreground mx-1">/</span>
                      <span className="text-chart-2">{sub.twpReplyCount}</span>
                    </TableCell>
                    <TableCell className="text-right text-sm">
                      {sub.lastMessageTime ? (
                        <div className="text-foreground" title={new Date(sub.lastMessageTime).toLocaleString()}>
                          {format(new Date(sub.lastMessageTime), 'MMM d, HH:mm')}
                        </div>
                      ) : (
                        <span className="text-muted-foreground italic">Never</span>
                      )}
                    </TableCell>
                  </TableRow>
                ))
              )}
            </TableBody>
          </Table>
        </div>
        {!isLoading && filteredSubscribers.length > 0 && (
          <div className="p-4 border-t border-border bg-muted/20 text-xs text-muted-foreground flex justify-between items-center">
            <span>Showing {filteredSubscribers.length} subscribers</span>
          </div>
        )}
      </div>
    </div>
  );
}
