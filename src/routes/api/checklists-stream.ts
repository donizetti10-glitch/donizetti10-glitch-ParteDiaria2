import { createFileRoute } from "@tanstack/react-router";
import {
  addChecklistsListener,
  removeChecklistsListener,
  getLatestServerChecklists,
} from "./checklists";
import type { EquipmentChecklist } from "@/lib/checklist";

export const Route = createFileRoute("/api/checklists-stream")({
  server: {
    handlers: {
      GET: async () => {
        let listener: (checklists: EquipmentChecklist[]) => void;
        const stream = new ReadableStream({
          start(controller) {
            const encoder = new TextEncoder();

            // Send initial checklists state immediately
            const current = getLatestServerChecklists();
            controller.enqueue(
              encoder.encode(`data: ${JSON.stringify({ type: "init", checklists: current })}\n\n`),
            );

            listener = (checklists: EquipmentChecklist[]) => {
              try {
                controller.enqueue(
                  encoder.encode(`data: ${JSON.stringify({ type: "update", checklists })}\n\n`),
                );
              } catch (e) {
                console.warn("SSE checklist stream enqueue error:", e);
              }
            };
            addChecklistsListener(listener);
          },
          cancel() {
            if (listener) removeChecklistsListener(listener);
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
