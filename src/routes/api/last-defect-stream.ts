import { createFileRoute } from "@tanstack/react-router";
import {
  addDefectListener,
  removeDefectListener,
  getLatestRecord,
  DefectRecord,
} from "./last-defect";

export const Route = createFileRoute("/api/last-defect-stream")({
  server: {
    handlers: {
      GET: async () => {
        let listener: (rec: DefectRecord) => void;
        const stream = new ReadableStream({
          start(controller) {
            const encoder = new TextEncoder();

            // Send initial record immediately upon connection
            const current = getLatestRecord();
            controller.enqueue(encoder.encode(`data: ${JSON.stringify(current)}\n\n`));

            listener = (rec: DefectRecord) => {
              try {
                controller.enqueue(encoder.encode(`data: ${JSON.stringify(rec)}\n\n`));
              } catch {}
            };
            addDefectListener(listener);
          },
          cancel() {
            if (listener) removeDefectListener(listener);
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
