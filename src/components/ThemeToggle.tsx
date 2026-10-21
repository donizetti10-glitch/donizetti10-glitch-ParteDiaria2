import { useEffect, useState } from "react";
import { Moon, Sun } from "lucide-react";

export function ThemeToggle({
  className = "",
  showLabel = true,
}: {
  className?: string;
  showLabel?: boolean;
}) {
  const [theme, setTheme] = useState<"light" | "dark">("light");

  useEffect(() => {
    // Check initial theme from localStorage or document class or system preference
    const storedTheme = localStorage.getItem("theme");
    if (storedTheme === "dark" || storedTheme === "light") {
      setTheme(storedTheme);
      document.documentElement.classList.toggle("dark", storedTheme === "dark");
    } else {
      const systemDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
      const initial = systemDark ? "dark" : "light";
      setTheme(initial);
      document.documentElement.classList.toggle("dark", systemDark);
    }
  }, []);

  const toggleTheme = () => {
    const nextTheme = theme === "dark" ? "light" : "dark";
    setTheme(nextTheme);
    localStorage.setItem("theme", nextTheme);
    document.documentElement.classList.toggle("dark", nextTheme === "dark");
  };

  const isDark = theme === "dark";

  return (
    <button
      type="button"
      onClick={toggleTheme}
      className={`btn-secondary text-xs flex items-center gap-1.5 transition-colors ${className}`}
      title={isDark ? "Mudar para Modo Claro" : "Mudar para Modo Escuro"}
      aria-label={isDark ? "Mudar para Modo Claro" : "Mudar para Modo Escuro"}
    >
      {isDark ? (
        <>
          <Sun className="h-4 w-4 text-amber-400 shrink-0" />
          {showLabel && <span>Modo Claro</span>}
        </>
      ) : (
        <>
          <Moon className="h-4 w-4 text-slate-700 dark:text-slate-300 shrink-0" />
          {showLabel && <span>Modo Escuro</span>}
        </>
      )}
    </button>
  );
}
