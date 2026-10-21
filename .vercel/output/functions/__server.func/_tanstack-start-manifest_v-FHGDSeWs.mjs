//#region node_modules/.nitro/vite/services/ssr/assets/_tanstack-start-manifest_v-FHGDSeWs.js
var tsrStartManifest = () => ({
  routes: {
    __root__: {
      filePath: "/app/applet/src/routes/__root.tsx",
      children: [
        "/",
        "/checklist",
        "/dashboard",
        "/sitemap.xml",
        "/api/checklists",
        "/api/checklists-stream",
        "/api/cloud-reports",
        "/api/cloud-reports-stream",
        "/api/last-defect",
        "/api/last-defect-stream",
      ],
      preloads: ["/assets/index-n8KTkd9n.js", "/assets/rolldown-runtime-QTnfLwEv.js"],
      scripts: [
        {
          attrs: {
            type: "module",
            async: !0,
            src: "/assets/index-n8KTkd9n.js",
          },
        },
      ],
    },
    "/": {
      filePath: "/app/applet/src/routes/index.tsx",
      children: void 0,
      preloads: [
        "/assets/routes-cAyyxGbL.js",
        "/assets/helpers-nemQBT-g.js",
        "/assets/FormControlBuilderCommon-2HvmNkMR.js",
      ],
    },
  },
});
//#endregion
export { tsrStartManifest };
