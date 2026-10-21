import { createFileRoute } from "@tanstack/react-router";
import { addReportsListener, removeReportsListener, getLatestServerReports } from "./cloud-reports";
import type { SavedCloudReport } from "@/lib/cloud-reports";

export const Route = createFileRoute("/api/cloud-reports-stream")({
  server: {
    handlers: {
      GET: async () => {
        let listener: (reports: SavedCloudReport[]) => void;
        const stream = new ReadableStream({
          start(controller) {
            const encoder = new TextEncoder();

            // Send initial reports state immediately upon connection
            const current = getLatestServerReports();
            controller.enqueue(
              encoder.encode(`data: ${JSON.stringify({ type: "init", reports: current })}\n\n`),
            );

            listener = (reports: SavedCloudReport[]) => {
              try {
                controller.enqueue(
                  encoder.encode(`data: ${JSON.stringify({ type: "update", reports })}\n\n`),
                );
              } catch (e) {
                console.warn("SSE stream enqueue error:", e);
              }
            };
            addReportsListener(listener);
          },
          cancel() {
            if (listener) removeReportsListener(listener);
          },
        });

        return new Response(stream, {
          headers: {
            "Content-Type": "text/event-stream",
            "Cache-Control": "no-cache",
            Connection: "keep-alive",
          },
        });
      },
    },
  },
});
