import { useGetDashboardSummary, useGetLabelStats, useGetAgentStats, useGetSequenceStats } from "@workspace/api-client-react";
import { useCredentials } from "@/contexts/CredentialsContext";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Skeleton } from "@/components/ui/skeleton";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer, PieChart, Pie, Cell } from "recharts";
import { Users, Activity, Moon, TrendingUp, AlertCircle, ArrowRight } from "lucide-react";
import { Link } from "wouter";

export default function Dashboard() {
  const { credentials } = useCredentials();
  
  const queryOpts = {
    apiToken: credentials?.apiToken || "",
    phoneNumberId: credentials?.phoneNumberId || ""
  };

  const { data: summary, isLoading: isLoadingSummary } = useGetDashboardSummary(queryOpts, {
    query: { enabled: !!credentials?.apiToken }
  });

  const { data: labelsData, isLoading: isLoadingLabels } = useGetLabelStats(queryOpts, {
    query: { enabled: !!credentials?.apiToken }
  });

  const { data: agentsData, isLoading: isLoadingAgents } = useGetAgentStats(queryOpts, {
    query: { enabled: !!credentials?.apiToken }
  });

  const { data: sequencesData, isLoading: isLoadingSequences } = useGetSequenceStats(queryOpts, {
    query: { enabled: !!credentials?.apiToken }
  });

  const PIE_COLORS = ["hsl(var(--chart-1))", "hsl(var(--chart-2))"];

  const topSequence = sequencesData?.sequences
    ? [...sequencesData.sequences].sort((a, b) => b.reactivationRate - a.reactivationRate)[0]
    : null;

  return (
    <div className="space-y-6 max-w-7xl mx-auto">
      <div>
        <h1 className="text-3xl font-bold tracking-tight text-foreground">Operations Overview</h1>
        <p className="text-muted-foreground mt-1">High-level metrics and performance highlights.</p>
      </div>

      {/* KPI Row */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <KpiCard
          title="Total Leads"
          value={summary?.totalLeads}
          icon={<Users className="w-4 h-4 text-muted-foreground" />}
          isLoading={isLoadingSummary}
        />
        <KpiCard
          title="Active Leads"
          value={summary?.activeLeads}
          icon={<Activity className="w-4 h-4 text-chart-2" />}
          isLoading={isLoadingSummary}
        />
        <KpiCard
          title="Dormant Leads"
          value={summary?.dormantLeads}
          icon={<Moon className="w-4 h-4 text-destructive" />}
          isLoading={isLoadingSummary}
        />
        <KpiCard
          title="Reactivation Rate"
          value={summary?.reactivationRate ? `${(summary.reactivationRate * 100).toFixed(1)}%` : "0%"}
          icon={<TrendingUp className="w-4 h-4 text-chart-1" />}
          isLoading={isLoadingSummary}
        />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Reply Volume Comparison */}
        <Card className="lg:col-span-1 border-border shadow-sm">
          <CardHeader>
            <CardTitle className="text-lg">Reply Volume</CardTitle>
            <CardDescription>User vs TWP Messages</CardDescription>
          </CardHeader>
          <CardContent className="h-[250px] flex items-center justify-center">
            {isLoadingSummary ? (
              <Skeleton className="w-full h-full rounded-full" />
            ) : (
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={[
                      { name: "User Replies", value: summary?.totalUserReplies || 0 },
                      { name: "TWP Replies", value: summary?.totalTwpReplies || 0 }
                    ]}
                    cx="50%"
                    cy="50%"
                    innerRadius={60}
                    outerRadius={80}
                    paddingAngle={2}
                    dataKey="value"
                  >
                    {[0, 1].map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={PIE_COLORS[index % PIE_COLORS.length]} />
                    ))}
                  </Pie>
                  <RechartsTooltip 
                    contentStyle={{ backgroundColor: 'hsl(var(--card))', borderColor: 'hsl(var(--border))' }}
                    itemStyle={{ color: 'hsl(var(--foreground))' }}
                  />
                </PieChart>
              </ResponsiveContainer>
            )}
          </CardContent>
        </Card>

        {/* Label Distribution */}
        <Card className="lg:col-span-2 border-border shadow-sm">
          <CardHeader>
            <CardTitle className="text-lg">Lead Distribution by Label</CardTitle>
            <CardDescription>Total vs Dormant count per pipeline stage</CardDescription>
          </CardHeader>
          <CardContent className="h-[250px]">
            {isLoadingLabels ? (
              <Skeleton className="w-full h-full" />
            ) : (
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={labelsData?.labels || []} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="hsl(var(--border))" />
                  <XAxis dataKey="labelName" tick={{ fill: 'hsl(var(--muted-foreground))', fontSize: 12 }} tickLine={false} axisLine={false} />
                  <YAxis tick={{ fill: 'hsl(var(--muted-foreground))', fontSize: 12 }} tickLine={false} axisLine={false} />
                  <RechartsTooltip 
                    cursor={{ fill: 'hsl(var(--muted)/0.5)' }}
                    contentStyle={{ backgroundColor: 'hsl(var(--card))', borderColor: 'hsl(var(--border))', borderRadius: '6px' }}
                  />
                  <Bar dataKey="count" name="Total" fill="hsl(var(--chart-2))" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="dormantCount" name="Dormant" fill="hsl(var(--destructive))" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            )}
          </CardContent>
        </Card>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Top Agent Snippet */}
        <Card className="border-border shadow-sm flex flex-col">
          <CardHeader className="flex flex-row items-center justify-between pb-2">
            <div className="space-y-1">
              <CardTitle className="text-lg">Agent Performance</CardTitle>
              <CardDescription>Top active agents</CardDescription>
            </div>
            <Link href="/agents" className="text-sm text-primary hover:underline flex items-center gap-1">
              View all <ArrowRight className="w-3 h-3" />
            </Link>
          </CardHeader>
          <CardContent className="flex-1">
            {isLoadingAgents ? (
              <div className="space-y-3">
                <Skeleton className="h-10 w-full" />
                <Skeleton className="h-10 w-full" />
                <Skeleton className="h-10 w-full" />
              </div>
            ) : (
              <div className="space-y-4 mt-2">
                {agentsData?.agents?.slice(0, 4).map(agent => (
                  <div key={agent.agentName} className="flex items-center justify-between">
                    <div>
                      <p className="font-medium text-sm">{agent.agentName}</p>
                      <p className="text-xs text-muted-foreground">{agent.leadsAssigned} assigned</p>
                    </div>
                    <div className="text-right">
                      <p className="text-sm font-bold text-foreground">{agent.activeLeads} active</p>
                      <p className="text-xs text-destructive">{agent.dormantLeads} dormant</p>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </CardContent>
        </Card>

        {/* Top Sequence Snippet */}
        <Card className="border-border shadow-sm bg-primary/5 border-primary/20">
          <CardHeader>
            <div className="flex items-center gap-2 mb-2">
              <AlertCircle className="w-5 h-5 text-primary" />
              <CardTitle className="text-lg text-primary">Top Performing Sequence</CardTitle>
            </div>
            <CardDescription>Highest reactivation rate</CardDescription>
          </CardHeader>
          <CardContent>
            {isLoadingSequences ? (
              <Skeleton className="h-24 w-full" />
            ) : topSequence ? (
              <div className="space-y-4">
                <div className="text-2xl font-bold text-foreground">{topSequence.sequenceName}</div>
                <div className="grid grid-cols-3 gap-4 border-t border-primary/10 pt-4">
                  <div>
                    <p className="text-xs text-muted-foreground uppercase tracking-wider mb-1">Sent</p>
                    <p className="text-lg font-medium">{topSequence.totalSent}</p>
                  </div>
                  <div>
                    <p className="text-xs text-muted-foreground uppercase tracking-wider mb-1">Replies</p>
                    <p className="text-lg font-medium">{topSequence.repliesAfterSequence}</p>
                  </div>
                  <div>
                    <p className="text-xs text-muted-foreground uppercase tracking-wider mb-1">Conv. Rate</p>
                    <p className="text-lg font-bold text-primary">{(topSequence.reactivationRate * 100).toFixed(1)}%</p>
                  </div>
                </div>
              </div>
            ) : (
              <div className="text-sm text-muted-foreground">No sequence data available.</div>
            )}
          </CardContent>
        </Card>
      </div>
    </div>
  );
}

function KpiCard({ title, value, icon, isLoading }: { title: string; value?: string | number; icon: React.ReactNode; isLoading: boolean }) {
  return (
    <Card className="border-border shadow-sm">
      <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
        <CardTitle className="text-sm font-medium text-muted-foreground">
          {title}
        </CardTitle>
        {icon}
      </CardHeader>
      <CardContent>
        {isLoading ? (
          <Skeleton className="h-8 w-20" />
        ) : (
          <div className="text-2xl font-bold tracking-tight text-foreground">{value ?? 0}</div>
        )}
      </CardContent>
    </Card>
  );
}
