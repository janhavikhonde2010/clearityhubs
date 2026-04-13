import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { Toaster } from "@/components/ui/toaster";
import { TooltipProvider } from "@/components/ui/tooltip";
import { CredentialsProvider } from "@/contexts/CredentialsContext";
import MainPage from "@/pages/MainPage";

const queryClient = new QueryClient({
  defaultOptions: {
    queries: {
      retry: 1,
      refetchOnWindowFocus: false,
      staleTime: 5 * 60 * 1000,
    },
  },
});

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <CredentialsProvider>
          <MainPage />
          <Toaster />
        </CredentialsProvider>
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;
