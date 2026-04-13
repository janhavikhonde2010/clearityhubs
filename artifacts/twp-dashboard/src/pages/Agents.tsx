import { useGetAgentStats } from "@workspace/api-client-react";
import { useCredentials } from "@/contexts/CredentialsContext";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Skeleton } from "@/components/ui/skeleton";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer, Legend } from "recharts";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Headset, Target, Clock, Zap } from "lucide-react";

export default function Agents() {
  const { credentials } = useCredentials();
  
  const queryOpts = {
    apiToken: credentials?.apiToken || "",
    phoneNumberId: credentials?.phoneNumberId || ""
  };

  const { data, isLoading } = useGetAgentStats(queryOpts, {
    query: { enabled: !!credentials?.apiToken }
  });

  const agents = data?.agents || [];
  
  // Sort agents by total assigned leads for the chart
  const chartData = [...agents].sort((a, b) => b.leadsAssigned - a.leadsAssigned);

  return (
    <div className="space-y-6 max-w-7xl mx-auto">
      <div className="flex items-center gap-3">
        <div className="w-10 h-10 rounded-lg bg-primary/10 flex items-center justify-center">
          <Headset className="w-5 h-5 text-primary" />
        </div>
        <div>
          <h1 className="text-3xl font-bold tracking-tight">Agent Performance</h1>
          <p className="text-muted-foreground mt-1">Workload distribution and response effectiveness by agent.</p>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <Card className="border-border shadow-sm">
          <CardHeader className="pb-2">
            <CardTitle className="text-sm font-medium text-muted-foreground flex items-center gap-2">
              <Target className="w-4 h-4" /> Total Agents
            </CardTitle>
          </CardHeader>
          <CardContent>
            {isLoading ? <Skeleton className="h-8 w-16" /> : (
              <div className="text-2xl font-bold">{agents.length}</div>
            )}
          </CardContent>
        </Card>
        <Card className="border-border shadow-sm">
          <CardHeader className="pb-2">
            <CardTitle className="text-sm font-medium text-muted-foreground flex items-center gap-2">
              <Clock className="w-4 h-4" /> Total Assigned
            </CardTitle>
          </CardHeader>
          <CardContent>
            {isLoading ? <Skeleton className="h-8 w-24" /> : (
              <div className="text-2xl font-bold text-primary">
                {agents.reduce((sum, a) => sum + a.leadsAssigned, 0).toLocaleString()}
              </div>
            )}
          </CardContent>
        </Card>
        <Card className="border-border shadow-sm">
          <CardHeader className="pb-2">
            <CardTitle className="text-sm font-medium text-muted-foreground flex items-center gap-2">
              <Zap className="w-4 h-4" /> Active Portfolio
            </CardTitle>
          </CardHeader>
          <CardContent>
            {isLoading ? <Skeleton className="h-8 w-24" /> : (
              <div className="text-2xl font-bold text-chart-2">
                {agents.reduce((sum, a) => sum + a.activeLeads, 0).toLocaleString()}
              </div>
            )}
          </CardContent>
        </Card>
        <Card className="border-border shadow-sm">
          <CardHeader className="pb-2">
            <CardTitle className="text-sm font-medium text-muted-foreground flex items-center gap-2">
              <Clock className="w-4 h-4" /> Dormant Risk
            </CardTitle>
          </CardHeader>
          <CardContent>
            {isLoading ? <Skeleton className="h-8 w-24" /> : (
              <div className="text-2xl font-bold text-destructive">
                {agents.reduce((sum, a) => sum + a.dormantLeads, 0).toLocaleString()}
              </div>
            )}
          </CardContent>
        </Card>
      </div>

      <Card className="border-border shadow-sm">
        <CardHeader>
          <CardTitle>Lead Workload Distribution</CardTitle>
          <CardDescription>Active vs Dormant leads assigned to each agent</CardDescription>
        </CardHeader>
        <CardContent className="h-[400px]">
          {isLoading ? (
            <Skeleton className="w-full h-full" />
          ) : agents.length === 0 ? (
            <div className="w-full h-full flex items-center justify-center text-muted-foreground">No agent data available</div>
          ) : (
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartData} margin={{ top: 20, right: 30, left: 0, bottom: 20 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="hsl(var(--border))" />
                <XAxis 
                  dataKey="agentName" 
                  tick={{ fill: 'hsl(var(--muted-foreground))', fontSize: 12 }} 
                  tickLine={false} 
                  axisLine={false} 
                  angle={-45}
                  textAnchor="end"
                  height={60}
                />
                <YAxis tick={{ fill: 'hsl(var(--muted-foreground))', fontSize: 12 }} tickLine={false} axisLine={false} />
                <RechartsTooltip 
                  cursor={{ fill: 'hsl(var(--muted)/0.5)' }}
                  contentStyle={{ backgroundColor: 'hsl(var(--card))', borderColor: 'hsl(var(--border))', borderRadius: '6px' }}
                  itemStyle={{ color: 'hsl(var(--foreground))' }}
                />
                <Legend wrapperStyle={{ paddingTop: '20px' }} />
                <Bar dataKey="activeLeads" name="Active Leads" stackId="a" fill="hsl(var(--primary))" radius={[0, 0, 4, 4]} />
                <Bar dataKey="dormantLeads" name="Dormant Leads" stackId="a" fill="hsl(var(--destructive))" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          )}
        </CardContent>
      </Card>

      <Card className="border-border shadow-sm">
        <CardHeader>
          <CardTitle>Detailed Agent Metrics</CardTitle>
        </CardHeader>
        <CardContent className="p-0">
          <div className="overflow-x-auto">
            <Table>
              <TableHeader className="bg-muted/50">
                <TableRow>
                  <TableHead>Agent Name</TableHead>
                  <TableHead className="text-right">Total Assigned</TableHead>
                  <TableHead className="text-right">Active</TableHead>
                  <TableHead className="text-right">Dormant</TableHead>
                  <TableHead className="text-right">Dormant %</TableHead>
                  <TableHead className="text-right">Avg User Replies</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {isLoading ? (
                  Array.from({ length: 5 }).map((_, i) => (
                    <TableRow key={i}>
                      <TableCell><Skeleton className="h-5 w-32" /></TableCell>
                      <TableCell className="text-right"><Skeleton className="h-5 w-12 ml-auto" /></TableCell>
                      <TableCell className="text-right"><Skeleton className="h-5 w-12 ml-auto" /></TableCell>
                      <TableCell className="text-right"><Skeleton className="h-5 w-12 ml-auto" /></TableCell>
                      <TableCell className="text-right"><Skeleton className="h-5 w-12 ml-auto" /></TableCell>
                      <TableCell className="text-right"><Skeleton className="h-5 w-12 ml-auto" /></TableCell>
                    </TableRow>
                  ))
                ) : agents.length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={6} className="h-32 text-center text-muted-foreground">
                      No agent data available.
                    </TableCell>
                  </TableRow>
                ) : (
                  [...agents].sort((a, b) => b.leadsAssigned - a.leadsAssigned).map((agent, idx) => {
                    const dormantPercent = agent.leadsAssigned > 0 
                      ? (agent.dormantLeads / agent.leadsAssigned) * 100 
                      : 0;
                      
                    return (
                      <TableRow key={agent.agentName} className={idx % 2 === 0 ? "bg-background" : "bg-muted/20"}>
                        <TableCell className="font-medium">{agent.agentName}</TableCell>
                        <TableCell className="text-right font-mono">{agent.leadsAssigned}</TableCell>
                        <TableCell className="text-right font-mono text-chart-2">{agent.activeLeads}</TableCell>
                        <TableCell className="text-right font-mono text-destructive">{agent.dormantLeads}</TableCell>
                        <TableCell className="text-right">
                          <span className={`px-2 py-1 rounded text-xs font-mono ${
                            dormantPercent > 50 ? 'bg-destructive/20 text-destructive' : 
                            dormantPercent > 30 ? 'bg-amber-500/20 text-amber-600 dark:text-amber-500' : 
                            'text-muted-foreground'
                          }`}>
                            {dormantPercent.toFixed(1)}%
                          </span>
                        </TableCell>
                        <TableCell className="text-right font-mono">{agent.avgUserReplies.toFixed(2)}</TableCell>
                      </TableRow>
                    );
                  })
                )}
              </TableBody>
            </Table>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
