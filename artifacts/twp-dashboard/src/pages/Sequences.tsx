import { useGetSequenceStats } from "@workspace/api-client-react";
import { useCredentials } from "@/contexts/CredentialsContext";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Skeleton } from "@/components/ui/skeleton";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer, Cell } from "recharts";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { MessageSquare, TrendingUp, AlertTriangle } from "lucide-react";

export default function Sequences() {
  const { credentials } = useCredentials();
  
  const queryOpts = {
    apiToken: credentials?.apiToken || "",
    phoneNumberId: credentials?.phoneNumberId || ""
  };

  const { data, isLoading } = useGetSequenceStats(queryOpts, {
    query: { enabled: !!credentials?.apiToken }
  });

  const sequences = data?.sequences || [];
  
  const sortedSequences = [...sequences].sort((a, b) => b.reactivationRate - a.reactivationRate);
  const bestSequence = sortedSequences.length > 0 ? sortedSequences[0] : null;
  const worstSequence = sortedSequences.length > 0 ? sortedSequences[sortedSequences.length - 1] : null;

  return (
    <div className="space-y-6 max-w-7xl mx-auto">
      <div className="flex items-center gap-3">
        <div className="w-10 h-10 rounded-lg bg-primary/10 flex items-center justify-center">
          <MessageSquare className="w-5 h-5 text-primary" />
        </div>
        <div>
          <h1 className="text-3xl font-bold tracking-tight">Sequence Analytics</h1>
          <p className="text-muted-foreground mt-1">Evaluate the effectiveness of automated marketing messages.</p>
        </div>
      </div>

      {bestSequence && (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <Card className="border-primary/30 shadow-md bg-gradient-to-br from-primary/10 to-background relative overflow-hidden">
            <div className="absolute top-0 right-0 w-32 h-32 bg-primary/10 rounded-bl-full -z-10" />
            <CardHeader className="pb-2">
              <CardTitle className="text-sm font-medium text-primary flex items-center gap-2">
                <TrendingUp className="w-4 h-4" /> Top Performing Sequence
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold mb-4">{bestSequence.sequenceName}</div>
              <div className="grid grid-cols-3 gap-4">
                <div>
                  <p className="text-xs text-muted-foreground mb-1">Conversion Rate</p>
                  <p className="text-xl font-bold text-primary">{(bestSequence.reactivationRate * 100).toFixed(1)}%</p>
                </div>
                <div>
                  <p className="text-xs text-muted-foreground mb-1">Total Sent</p>
                  <p className="text-lg font-mono">{bestSequence.totalSent.toLocaleString()}</p>
                </div>
                <div>
                  <p className="text-xs text-muted-foreground mb-1">Replies</p>
                  <p className="text-lg font-mono">{bestSequence.repliesAfterSequence.toLocaleString()}</p>
                </div>
              </div>
            </CardContent>
          </Card>

          {worstSequence && bestSequence.sequenceName !== worstSequence.sequenceName && (
            <Card className="border-destructive/30 shadow-sm bg-destructive/5 relative overflow-hidden">
              <div className="absolute top-0 right-0 w-32 h-32 bg-destructive/5 rounded-bl-full -z-10" />
              <CardHeader className="pb-2">
                <CardTitle className="text-sm font-medium text-destructive flex items-center gap-2">
                  <AlertTriangle className="w-4 h-4" /> Needs Improvement
                </CardTitle>
              </CardHeader>
              <CardContent>
                <div className="text-2xl font-bold mb-4">{worstSequence.sequenceName}</div>
                <div className="grid grid-cols-3 gap-4">
                  <div>
                    <p className="text-xs text-muted-foreground mb-1">Conversion Rate</p>
                    <p className="text-xl font-bold text-destructive">{(worstSequence.reactivationRate * 100).toFixed(1)}%</p>
                  </div>
                  <div>
                    <p className="text-xs text-muted-foreground mb-1">Total Sent</p>
                    <p className="text-lg font-mono">{worstSequence.totalSent.toLocaleString()}</p>
                  </div>
                  <div>
                    <p className="text-xs text-muted-foreground mb-1">Replies</p>
                    <p className="text-lg font-mono">{worstSequence.repliesAfterSequence.toLocaleString()}</p>
                  </div>
                </div>
              </CardContent>
            </Card>
          )}
        </div>
      )}

      <Card className="border-border shadow-sm">
        <CardHeader>
          <CardTitle>Sequence Conversion Rates</CardTitle>
          <CardDescription>Percentage of leads reactivated after receiving each sequence</CardDescription>
        </CardHeader>
        <CardContent className="h-[350px]">
          {isLoading ? (
            <Skeleton className="w-full h-full" />
          ) : sequences.length === 0 ? (
            <div className="w-full h-full flex items-center justify-center text-muted-foreground">No sequence data available</div>
          ) : (
            <ResponsiveContainer width="100%" height="100%">
              <BarChart 
                data={sortedSequences} 
                layout="vertical" 
                margin={{ top: 5, right: 30, left: 20, bottom: 5 }}
              >
                <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="hsl(var(--border))" />
                <XAxis 
                  type="number" 
                  domain={[0, 1]} 
                  tickFormatter={(val) => `${(val * 100).toFixed(0)}%`}
                  tick={{ fill: 'hsl(var(--muted-foreground))', fontSize: 12 }} 
                  tickLine={false} 
                  axisLine={false} 
                />
                <YAxis 
                  dataKey="sequenceName" 
                  type="category" 
                  tick={{ fill: 'hsl(var(--foreground))', fontSize: 12, fontWeight: 500 }} 
                  tickLine={false} 
                  axisLine={false}
                  width={150}
                />
                <RechartsTooltip 
                  cursor={{ fill: 'hsl(var(--muted)/0.5)' }}
                  contentStyle={{ backgroundColor: 'hsl(var(--card))', borderColor: 'hsl(var(--border))', borderRadius: '6px' }}
                  formatter={(value: number) => [`${(value * 100).toFixed(2)}%`, 'Reactivation Rate']}
                />
                <Bar dataKey="reactivationRate" radius={[0, 4, 4, 0]}>
                  {sortedSequences.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={index === 0 ? 'hsl(var(--primary))' : 'hsl(var(--chart-2))'} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          )}
        </CardContent>
      </Card>

      <Card className="border-border shadow-sm">
        <CardHeader>
          <CardTitle>All Sequences</CardTitle>
        </CardHeader>
        <CardContent className="p-0">
          <div className="overflow-x-auto">
            <Table>
              <TableHeader className="bg-muted/50">
                <TableRow>
                  <TableHead>Sequence Name</TableHead>
                  <TableHead className="text-right">Total Sent</TableHead>
                  <TableHead className="text-right">Replies After</TableHead>
                  <TableHead className="text-right">Reactivation Rate</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {isLoading ? (
                  Array.from({ length: 5 }).map((_, i) => (
                    <TableRow key={i}>
                      <TableCell><Skeleton className="h-5 w-48" /></TableCell>
                      <TableCell className="text-right"><Skeleton className="h-5 w-16 ml-auto" /></TableCell>
                      <TableCell className="text-right"><Skeleton className="h-5 w-16 ml-auto" /></TableCell>
                      <TableCell className="text-right"><Skeleton className="h-5 w-16 ml-auto" /></TableCell>
                    </TableRow>
                  ))
                ) : sortedSequences.length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={4} className="h-32 text-center text-muted-foreground">
                      No sequence data available.
                    </TableCell>
                  </TableRow>
                ) : (
                  sortedSequences.map((seq, idx) => (
                    <TableRow key={seq.sequenceName} className={idx === 0 ? "bg-primary/5 hover:bg-primary/10" : ""}>
                      <TableCell className="font-medium">
                        {seq.sequenceName}
                        {idx === 0 && <span className="ml-2 text-[10px] uppercase tracking-wider bg-primary text-primary-foreground px-2 py-0.5 rounded-full">Top</span>}
                      </TableCell>
                      <TableCell className="text-right font-mono">{seq.totalSent.toLocaleString()}</TableCell>
                      <TableCell className="text-right font-mono">{seq.repliesAfterSequence.toLocaleString()}</TableCell>
                      <TableCell className="text-right font-mono font-bold text-primary">
                        {(seq.reactivationRate * 100).toFixed(2)}%
                      </TableCell>
                    </TableRow>
                  ))
                )}
              </TableBody>
            </Table>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
