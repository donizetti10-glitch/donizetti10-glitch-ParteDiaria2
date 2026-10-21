import { useState, useMemo } from "react";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  CartesianGrid,
  Legend,
  PieChart,
  Pie,
  Cell,
} from "recharts";
import {
  CheckCircle2,
  XCircle,
  AlertOctagon,
  FileQuestion,
  TrendingUp,
  Target,
  Ruler,
  Users,
  Calendar,
  Layers,
  ArrowUpRight,
} from "lucide-react";
import { type SavedCloudReport } from "@/lib/cloud-reports";
import { calculateGdaAnalytics, GDA_STATUS_CONFIG, type GdaStatusKey } from "@/lib/gda-analytics";

interface GdaAnalyticsDashboardProps {
  reports: SavedCloudReport[];
  selectedGdaFilter?: GdaStatusKey | "todos";
  onSelectGdaFilter?: (status: GdaStatusKey | "todos") => void;
}

export function GdaAnalyticsDashboard({
  reports,
  selectedGdaFilter = "todos",
  onSelectGdaFilter,
}: GdaAnalyticsDashboardProps) {
  const [activeTab, setActiveTab] = useState<"status" | "timeline" | "responsible" | "meters">(
    "status",
  );

  const analytics = useMemo(() => calculateGdaAnalytics(reports), [reports]);

  return (
    <div className="space-y-6">
      {/* SEÇÃO PRINCIPAL DE CARDS E INDICADORES DE GDA & PLANEJADO VS EXECUTADO */}
      <div className="space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div>
            <h3 className="text-sm font-bold uppercase tracking-wider text-foreground flex items-center gap-2">
              <Target className="h-4 w-4 text-primary" />
              <span>Controle de GDAs e Atividades: Planejadas vs. Executadas</span>
            </h3>
            <p className="text-xs text-muted-foreground">
              Acompanhamento de emissão de GDA, cumprimento do plano operacional, cancelamentos e
              reprovações.
            </p>
          </div>
          {onSelectGdaFilter && selectedGdaFilter !== "todos" && (
            <button
              onClick={() => onSelectGdaFilter("todos")}
              className="text-xs font-semibold text-primary hover:underline inline-flex items-center gap-1"
            >
              <span>Ver todos os status</span>
            </button>
          )}
        </div>

        {/* CARDS DE RESUMO KPI */}
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
          {/* Card 1: Planejadas vs Executadas */}
          <div
            onClick={() => onSelectGdaFilter?.("todos")}
            className={`cursor-pointer rounded-xl border p-3.5 transition-all relative overflow-hidden ${
              selectedGdaFilter === "todos"
                ? "border-primary bg-primary/5 shadow-xs"
                : "border-border bg-card hover:border-border/80"
            }`}
          >
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                Planejadas
              </span>
              <Target className="h-4 w-4 text-blue-500" />
            </div>
            <div className="mt-2 flex items-baseline gap-1.5">
              <span className="font-mono text-2xl font-bold text-foreground">
                {analytics.totalPlanned}
              </span>
              <span className="text-[11px] text-muted-foreground">atividades</span>
            </div>
            <div className="mt-1 text-[10px] text-muted-foreground flex items-center justify-between">
              <span>Executadas: {analytics.totalExecuted}</span>
              <span className="font-bold text-emerald-600 dark:text-emerald-400">
                {analytics.adherenceRate}%
              </span>
            </div>
            <div className="mt-2 w-full bg-secondary rounded-full h-1.5 overflow-hidden">
              <div
                className="bg-emerald-500 h-1.5 rounded-full transition-all duration-500"
                style={{ width: `${Math.min(100, analytics.adherenceRate)}%` }}
              />
            </div>
          </div>

          {/* Card 2: GDA Executada */}
          <div
            onClick={() => onSelectGdaFilter?.("executada")}
            className={`cursor-pointer rounded-xl border p-3.5 transition-all ${
              selectedGdaFilter === "executada"
                ? "border-emerald-500 bg-emerald-500/10 shadow-xs ring-1 ring-emerald-500"
                : "border-border bg-card hover:border-emerald-500/40"
            }`}
          >
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-semibold uppercase tracking-wider text-emerald-700 dark:text-emerald-400">
                GDA Executada
              </span>
              <CheckCircle2 className="h-4 w-4 text-emerald-500" />
            </div>
            <div className="mt-2 flex items-baseline gap-1.5">
              <span className="font-mono text-2xl font-bold text-emerald-600 dark:text-emerald-400">
                {analytics.statusDistribution.find((s) => s.key === "executada")?.value || 0}
              </span>
              <span className="text-[11px] text-muted-foreground">
                ({analytics.statusDistribution.find((s) => s.key === "executada")?.percentage || 0}
                %)
              </span>
            </div>
            <p className="mt-1 text-[10px] text-muted-foreground truncate">
              Concluídas com produção
            </p>
          </div>

          {/* Card 3: GDA Cancelada */}
          <div
            onClick={() => onSelectGdaFilter?.("cancelada")}
            className={`cursor-pointer rounded-xl border p-3.5 transition-all ${
              selectedGdaFilter === "cancelada"
                ? "border-rose-500 bg-rose-500/10 shadow-xs ring-1 ring-rose-500"
                : "border-border bg-card hover:border-rose-500/40"
            }`}
          >
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-semibold uppercase tracking-wider text-rose-700 dark:text-rose-400">
                GDA Cancelada
              </span>
              <XCircle className="h-4 w-4 text-rose-500" />
            </div>
            <div className="mt-2 flex items-baseline gap-1.5">
              <span className="font-mono text-2xl font-bold text-rose-600 dark:text-rose-400">
                {analytics.totalCancelled}
              </span>
              <span className="text-[11px] text-muted-foreground">
                ({analytics.statusDistribution.find((s) => s.key === "cancelada")?.percentage || 0}
                %)
              </span>
            </div>
            <p className="mt-1 text-[10px] text-muted-foreground truncate">
              Interrupção / Suspensão
            </p>
          </div>

          {/* Card 4: Sem GDA */}
          <div
            onClick={() => onSelectGdaFilter?.("sem_gda")}
            className={`cursor-pointer rounded-xl border p-3.5 transition-all ${
              selectedGdaFilter === "sem_gda"
                ? "border-slate-500 bg-slate-500/10 shadow-xs ring-1 ring-slate-500"
                : "border-border bg-card hover:border-slate-500/40"
            }`}
          >
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-700 dark:text-slate-300">
                Sem GDA
              </span>
              <FileQuestion className="h-4 w-4 text-slate-500" />
            </div>
            <div className="mt-2 flex items-baseline gap-1.5">
              <span className="font-mono text-2xl font-bold text-slate-600 dark:text-slate-300">
                {analytics.totalSemGda}
              </span>
              <span className="text-[11px] text-muted-foreground">
                ({analytics.statusDistribution.find((s) => s.key === "sem_gda")?.percentage || 0}%)
              </span>
            </div>
            <p className="mt-1 text-[10px] text-muted-foreground truncate">
              Sem numeração informada
            </p>
          </div>

          {/* Card 5: GDA Reprovada */}
          <div
            onClick={() => onSelectGdaFilter?.("reprovada")}
            className={`cursor-pointer rounded-xl border p-3.5 transition-all ${
              selectedGdaFilter === "reprovada"
                ? "border-amber-500 bg-amber-500/10 shadow-xs ring-1 ring-amber-500"
                : "border-border bg-card hover:border-amber-500/40"
            }`}
          >
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-semibold uppercase tracking-wider text-amber-700 dark:text-amber-400">
                GDA Reprovada
              </span>
              <AlertOctagon className="h-4 w-4 text-amber-500" />
            </div>
            <div className="mt-2 flex items-baseline gap-1.5">
              <span className="font-mono text-2xl font-bold text-amber-600 dark:text-amber-400">
                {analytics.totalReproved}
              </span>
              <span className="text-[11px] text-muted-foreground">
                ({analytics.statusDistribution.find((s) => s.key === "reprovada")?.percentage || 0}
                %)
              </span>
            </div>
            <p className="mt-1 text-[10px] text-muted-foreground truncate">
              Com defeitos / Não conforme
            </p>
          </div>

          {/* Card 6: Execução Parcial / Metragem */}
          <div
            onClick={() => onSelectGdaFilter?.("parcial")}
            className={`cursor-pointer rounded-xl border p-3.5 transition-all ${
              selectedGdaFilter === "parcial"
                ? "border-sky-500 bg-sky-500/10 shadow-xs ring-1 ring-sky-500"
                : "border-border bg-card hover:border-sky-500/40"
            }`}
          >
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-semibold uppercase tracking-wider text-sky-700 dark:text-sky-400">
                Exec. Parcial
              </span>
              <TrendingUp className="h-4 w-4 text-sky-500" />
            </div>
            <div className="mt-2 flex items-baseline gap-1.5">
              <span className="font-mono text-2xl font-bold text-sky-600 dark:text-sky-400">
                {analytics.totalPartial}
              </span>
              <span className="text-[11px] text-muted-foreground">
                ({analytics.statusDistribution.find((s) => s.key === "parcial")?.percentage || 0}%)
              </span>
            </div>
            <p className="mt-1 text-[10px] text-muted-foreground truncate">
              Abaixo da meta planejada
            </p>
          </div>
        </div>
      </div>

      {/* GRÁFICOS INTERATIVOS COM RECHARTS */}
      <div className="rounded-xl border border-border bg-card p-4 sm:p-5 shadow-sm space-y-4">
        {/* TAB CONTROLS DOS GRÁFICOS */}
        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-border/60 pb-3">
          <div className="flex flex-wrap items-center gap-1.5">
            <button
              onClick={() => setActiveTab("status")}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg transition-all ${
                activeTab === "status"
                  ? "bg-primary text-primary-foreground shadow-xs"
                  : "bg-secondary/60 text-muted-foreground hover:text-foreground hover:bg-secondary"
              }`}
            >
              <Layers className="h-3.5 w-3.5" />
              <span>Status das GDAs</span>
            </button>

            <button
              onClick={() => setActiveTab("timeline")}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg transition-all ${
                activeTab === "timeline"
                  ? "bg-primary text-primary-foreground shadow-xs"
                  : "bg-secondary/60 text-muted-foreground hover:text-foreground hover:bg-secondary"
              }`}
            >
              <Calendar className="h-3.5 w-3.5" />
              <span>Planejado vs. Executado por Dia</span>
            </button>

            <button
              onClick={() => setActiveTab("responsible")}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg transition-all ${
                activeTab === "responsible"
                  ? "bg-primary text-primary-foreground shadow-xs"
                  : "bg-secondary/60 text-muted-foreground hover:text-foreground hover:bg-secondary"
              }`}
            >
              <Users className="h-3.5 w-3.5" />
              <span>Aderência por Responsável</span>
            </button>

            <button
              onClick={() => setActiveTab("meters")}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg transition-all ${
                activeTab === "meters"
                  ? "bg-primary text-primary-foreground shadow-xs"
                  : "bg-secondary/60 text-muted-foreground hover:text-foreground hover:bg-secondary"
              }`}
            >
              <Ruler className="h-3.5 w-3.5" />
              <span>Metros: Programado vs. Executado</span>
            </button>
          </div>

          <div className="text-xs text-muted-foreground flex items-center gap-1.5">
            <ArrowUpRight className="h-3.5 w-3.5 text-primary" />
            <span>{reports.length} relatórios analisados</span>
          </div>
        </div>

        {/* TAB 1: STATUS DAS GDAs (DONUT & BARRAS LADO A LADO) */}
        {activeTab === "status" && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-center">
            {/* Donut Chart */}
            <div className="lg:col-span-6 h-72 w-full flex flex-col items-center justify-center">
              <h4 className="text-xs font-bold uppercase tracking-wider text-muted-foreground mb-2 text-center">
                Proporção das GDAs por Status Operacional
              </h4>
              {analytics.statusDistribution.length === 0 ? (
                <div className="h-56 flex items-center justify-center text-xs text-muted-foreground">
                  Nenhum dado disponível.
                </div>
              ) : (
                <ResponsiveContainer width="100%" height="90%">
                  <PieChart>
                    <Pie
                      data={analytics.statusDistribution}
                      cx="50%"
                      cy="50%"
                      innerRadius={60}
                      outerRadius={95}
                      paddingAngle={4}
                      dataKey="value"
                    >
                      {analytics.statusDistribution.map((entry) => (
                        <Cell key={`cell-${entry.key}`} fill={entry.color} />
                      ))}
                    </Pie>
                    <Tooltip
                      formatter={(val: unknown, name: unknown) => [
                        `${String(val ?? 0)} GDAs`,
                        String(name),
                      ]}
                      contentStyle={{
                        borderRadius: "8px",
                        fontSize: "12px",
                        boxShadow: "0 4px 6px -1px rgb(0 0 0 / 0.1)",
                      }}
                    />
                    <Legend wrapperStyle={{ fontSize: 11, paddingTop: 10 }} />
                  </PieChart>
                </ResponsiveContainer>
              )}
            </div>

            {/* Status Breakdown & Details */}
            <div className="lg:col-span-6 space-y-3">
              <h4 className="text-xs font-bold uppercase tracking-wider text-muted-foreground">
                Detalhamento dos Indicadores de GDA
              </h4>
              <div className="space-y-2">
                {(Object.keys(GDA_STATUS_CONFIG) as GdaStatusKey[]).map((key) => {
                  const cfg = GDA_STATUS_CONFIG[key];
                  const item = analytics.statusDistribution.find((s) => s.key === key);
                  const count = item?.value || 0;
                  const pct = item?.percentage || 0;

                  return (
                    <div
                      key={key}
                      onClick={() => onSelectGdaFilter?.(key)}
                      className={`p-2.5 rounded-lg border transition-all cursor-pointer flex items-center justify-between ${
                        selectedGdaFilter === key
                          ? "border-primary bg-primary/5"
                          : "border-border/60 bg-secondary/30 hover:bg-secondary/60"
                      }`}
                    >
                      <div className="flex items-center gap-2.5 min-w-0">
                        <span
                          className="h-3 w-3 rounded-full shrink-0"
                          style={{ backgroundColor: cfg.color }}
                        />
                        <div className="min-w-0">
                          <div className="text-xs font-semibold text-foreground truncate">
                            {cfg.label}
                          </div>
                          <div className="text-[10px] text-muted-foreground truncate">
                            {cfg.description}
                          </div>
                        </div>
                      </div>

                      <div className="text-right shrink-0 ml-3">
                        <span className="font-mono text-xs font-bold text-foreground">{count}</span>
                        <span className="text-[10px] text-muted-foreground ml-1.5">({pct}%)</span>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}

        {/* TAB 2: EVOLUÇÃO TEMPORAL - PLANEJADO VS EXECUTADO POR DIA */}
        {activeTab === "timeline" && (
          <div className="space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div>
                <h4 className="text-xs font-bold uppercase tracking-wider text-foreground">
                  Atividades Planejadas vs. Executadas por Data
                </h4>
                <p className="text-[11px] text-muted-foreground">
                  Comparação direta das tarefas programadas, concluídas e canceladas em cada
                  jornada.
                </p>
              </div>
              <div className="flex items-center gap-3 text-xs">
                <span className="flex items-center gap-1 text-blue-600 dark:text-blue-400 font-medium">
                  <span className="h-2.5 w-2.5 rounded-xs bg-blue-500 inline-block" /> Planejadas
                </span>
                <span className="flex items-center gap-1 text-emerald-600 dark:text-emerald-400 font-medium">
                  <span className="h-2.5 w-2.5 rounded-xs bg-emerald-500 inline-block" /> Executadas
                </span>
                <span className="flex items-center gap-1 text-rose-600 dark:text-rose-400 font-medium">
                  <span className="h-2.5 w-2.5 rounded-xs bg-rose-500 inline-block" /> Canceladas
                </span>
                <span className="flex items-center gap-1 text-amber-600 dark:text-amber-400 font-medium">
                  <span className="h-2.5 w-2.5 rounded-xs bg-amber-500 inline-block" /> Reprovadas
                </span>
              </div>
            </div>

            {analytics.dailyTimeline.length === 0 ? (
              <div className="h-64 flex items-center justify-center text-xs text-muted-foreground">
                Nenhum dado diário para o período selecionado.
              </div>
            ) : (
              <div className="h-72 w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart
                    data={analytics.dailyTimeline}
                    margin={{ top: 10, right: 10, left: -20, bottom: 20 }}
                  >
                    <CartesianGrid strokeDasharray="3 3" opacity={0.3} />
                    <XAxis
                      dataKey="dateLabel"
                      tick={{ fontSize: 10 }}
                      interval={0}
                      angle={-25}
                      textAnchor="end"
                    />
                    <YAxis tick={{ fontSize: 10 }} allowDecimals={false} />
                    <Tooltip
                      formatter={(val: unknown, name: unknown) => [
                        `${String(val ?? 0)} atividades`,
                        String(name),
                      ]}
                      labelFormatter={(label: unknown) => `Data: ${String(label)}`}
                      contentStyle={{
                        borderRadius: "8px",
                        fontSize: "12px",
                      }}
                    />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    <Bar
                      dataKey="planejadas"
                      name="Planejadas"
                      fill="#3b82f6"
                      radius={[4, 4, 0, 0]}
                    />
                    <Bar
                      dataKey="executadas"
                      name="Executadas"
                      fill="#16a34a"
                      radius={[4, 4, 0, 0]}
                    />
                    <Bar
                      dataKey="canceladas"
                      name="Canceladas"
                      fill="#dc2626"
                      radius={[4, 4, 0, 0]}
                    />
                    <Bar
                      dataKey="reprovadas"
                      name="Reprovadas"
                      fill="#ea580c"
                      radius={[4, 4, 0, 0]}
                    />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}
          </div>
        )}

        {/* TAB 3: ADERÊNCIA POR RESPONSÁVEL */}
        {activeTab === "responsible" && (
          <div className="space-y-3">
            <div className="flex items-center justify-between">
              <div>
                <h4 className="text-xs font-bold uppercase tracking-wider text-foreground">
                  Desempenho e Aderência por Encarregado / Responsável
                </h4>
                <p className="text-[11px] text-muted-foreground">
                  Percentual de cumprimento de atividades planejadas e volume de cancelamentos por
                  equipe.
                </p>
              </div>
            </div>

            {analytics.responsibleStats.length === 0 ? (
              <div className="h-64 flex items-center justify-center text-xs text-muted-foreground">
                Nenhum responsável registrado.
              </div>
            ) : (
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                <div className="lg:col-span-7 h-72 w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      data={analytics.responsibleStats.slice(0, 8)}
                      layout="vertical"
                      margin={{ top: 5, right: 20, left: 10, bottom: 5 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" opacity={0.3} />
                      <XAxis type="number" tick={{ fontSize: 10 }} allowDecimals={false} />
                      <YAxis dataKey="nome" type="category" tick={{ fontSize: 10 }} width={90} />
                      <Tooltip
                        formatter={(val: unknown, name: unknown) => [
                          `${String(val ?? 0)} atividades`,
                          String(name),
                        ]}
                        contentStyle={{
                          borderRadius: "8px",
                          fontSize: "12px",
                        }}
                      />
                      <Legend wrapperStyle={{ fontSize: 11 }} />
                      <Bar
                        dataKey="planejadas"
                        name="Planejadas"
                        fill="#3b82f6"
                        radius={[0, 4, 4, 0]}
                      />
                      <Bar
                        dataKey="executadas"
                        name="Executadas"
                        fill="#16a34a"
                        radius={[0, 4, 4, 0]}
                      />
                      <Bar
                        dataKey="canceladas"
                        name="Canceladas"
                        fill="#dc2626"
                        radius={[0, 4, 4, 0]}
                      />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div className="lg:col-span-5 space-y-2 overflow-y-auto max-h-72 pr-1">
                  {analytics.responsibleStats.map((resp, i) => (
                    <div
                      key={i}
                      className="p-2.5 rounded-lg border border-border/60 bg-secondary/30 text-xs space-y-1.5"
                    >
                      <div className="flex items-center justify-between">
                        <span className="font-semibold text-foreground truncate">{resp.nome}</span>
                        <span
                          className={`font-mono font-bold px-1.5 py-0.5 rounded text-[10px] ${
                            resp.taxaAderencia >= 80
                              ? "bg-emerald-500/20 text-emerald-700 dark:text-emerald-300"
                              : resp.taxaAderencia >= 50
                                ? "bg-amber-500/20 text-amber-700 dark:text-amber-300"
                                : "bg-rose-500/20 text-rose-700 dark:text-rose-300"
                          }`}
                        >
                          {resp.taxaAderencia}% Eficácia
                        </span>
                      </div>
                      <div className="grid grid-cols-4 gap-1 text-[10px] text-muted-foreground text-center">
                        <div className="bg-background/60 p-1 rounded">
                          <div className="font-bold text-foreground font-mono">
                            {resp.planejadas}
                          </div>
                          <div>Plan.</div>
                        </div>
                        <div className="bg-background/60 p-1 rounded">
                          <div className="font-bold text-emerald-600 font-mono">
                            {resp.executadas}
                          </div>
                          <div>Exec.</div>
                        </div>
                        <div className="bg-background/60 p-1 rounded">
                          <div className="font-bold text-rose-600 font-mono">{resp.canceladas}</div>
                          <div>Canc.</div>
                        </div>
                        <div className="bg-background/60 p-1 rounded">
                          <div className="font-bold text-slate-500 font-mono">{resp.semGda}</div>
                          <div>S/GDA</div>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* TAB 4: METROS - PROGRAMADO VS EXECUTADO */}
        {activeTab === "meters" && (
          <div className="space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div>
                <h4 className="text-xs font-bold uppercase tracking-wider text-foreground">
                  Extensão Linear: Metros Programados vs. Executados
                </h4>
                <p className="text-[11px] text-muted-foreground">
                  Metragem acumulada de produção programada na via férrea em comparação com os
                  metros efetivamente produzidos.
                </p>
              </div>

              {/* Indicador de Aderência Linear */}
              <div className="flex items-center gap-3 bg-secondary/60 px-3 py-1.5 rounded-lg border border-border">
                <div className="text-right">
                  <div className="text-[10px] text-muted-foreground uppercase font-semibold">
                    Cumprimento Físico
                  </div>
                  <div className="font-mono text-sm font-bold text-emerald-600 dark:text-emerald-400">
                    {analytics.totalExecutedMeters}m de {analytics.totalPlannedMeters}m (
                    {analytics.metersAdherenceRate}%)
                  </div>
                </div>
              </div>
            </div>

            {analytics.dailyTimeline.length === 0 ? (
              <div className="h-64 flex items-center justify-center text-xs text-muted-foreground">
                Nenhum registro de metragem encontrado.
              </div>
            ) : (
              <div className="h-72 w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart
                    data={analytics.dailyTimeline}
                    margin={{ top: 10, right: 10, left: 10, bottom: 20 }}
                  >
                    <CartesianGrid strokeDasharray="3 3" opacity={0.3} />
                    <XAxis
                      dataKey="dateLabel"
                      tick={{ fontSize: 10 }}
                      interval={0}
                      angle={-25}
                      textAnchor="end"
                    />
                    <YAxis tick={{ fontSize: 10 }} unit="m" />
                    <Tooltip
                      formatter={(val: unknown, name: unknown) => [
                        `${Number(val || 0).toLocaleString("pt-BR")} metros`,
                        String(name),
                      ]}
                      labelFormatter={(label: unknown) => `Data: ${String(label)}`}
                      contentStyle={{
                        borderRadius: "8px",
                        fontSize: "12px",
                      }}
                    />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    <Bar
                      dataKey="programadoMeters"
                      name="Programado (m)"
                      fill="#6366f1"
                      radius={[4, 4, 0, 0]}
                    />
                    <Bar
                      dataKey="executadoMeters"
                      name="Executado (m)"
                      fill="#10b981"
                      radius={[4, 4, 0, 0]}
                    />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
