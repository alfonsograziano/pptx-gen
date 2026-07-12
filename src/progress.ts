// A tiny zero-dependency progress reporter for deck builds.
//
// Goals:
// - Show every step as it happens, so a 30s build never feels stuck.
// - Render a live spinner + elapsed time on a TTY so the user sees motion.
// - Time each step and print a clean summary at the end.
// - Degrade gracefully when output is piped (no TTY): plain start/finish lines.

const SPINNER_FRAMES = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"];
const TICK_MS = 80;

type FinishedStep = {
  label: string;
  ms: number;
  detail?: string;
};

function color(code: string, text: string): string {
  if (!process.stdout.isTTY) return text;
  return `\x1b[${code}m${text}\x1b[0m`;
}

const dim = (text: string) => color("2", text);

/** Dim, parenthesised annotation for inline labels, e.g. "(custom)". */
export const dimKind = (text: string): string => dim(`(${text})`);
const cyan = (text: string) => color("36", text);
const green = (text: string) => color("32", text);
const bold = (text: string) => color("1", text);

function formatMs(ms: number): string {
  if (ms < 1000) return `${Math.round(ms)}ms`;
  return `${(ms / 1000).toFixed(1)}s`;
}

function now(): number {
  return Number(process.hrtime.bigint() / 1_000_000n);
}

function bar(fraction: number, width = 16): string {
  const clamped = Math.max(0, Math.min(1, fraction));
  const filled = Math.round(clamped * width);
  return `${"█".repeat(filled)}${"░".repeat(width - filled)}`;
}

export class BuildProgress {
  private readonly isTty = Boolean(process.stdout.isTTY);
  private readonly startedAt = now();
  private readonly finished: FinishedStep[] = [];
  private totalSteps: number;
  private currentLabel: string | null = null;
  private currentStartedAt = 0;
  private frame = 0;
  private timer: NodeJS.Timeout | null = null;

  constructor(title: string, expectedSteps: number) {
    this.totalSteps = Math.max(1, expectedSteps);
    process.stdout.write(`\n${bold("📦 " + title)}\n\n`);
  }

  /** Run `fn` as a named step, timing it and updating the live line. */
  async step<T>(label: string, fn: () => Promise<T> | T): Promise<T> {
    this.begin(label);
    try {
      const result = await fn();
      this.end();
      return result;
    } catch (error) {
      this.fail(error);
      throw error;
    }
  }

  /** Emit a one-off note above the live line (does not count as a step). */
  note(message: string): void {
    this.clearLine();
    process.stdout.write(`  ${dim("·")} ${dim(message)}\n`);
    this.render();
  }

  private begin(label: string): void {
    this.currentLabel = label;
    this.currentStartedAt = now();
    if (this.isTty) {
      this.frame = 0;
      this.render();
      this.timer = setInterval(() => this.render(), TICK_MS);
      this.timer.unref?.();
    } else {
      process.stdout.write(`  → ${label}\n`);
    }
  }

  private end(): void {
    if (this.currentLabel === null) return;
    const ms = now() - this.currentStartedAt;
    const label = this.currentLabel;
    this.stopTimer();
    this.clearLine();
    this.finished.push({ label, ms });
    process.stdout.write(`  ${green("✓")} ${label} ${dim(formatMs(ms))}\n`);
    this.currentLabel = null;
    this.render();
  }

  private fail(error: unknown): void {
    if (this.currentLabel === null) return;
    const ms = now() - this.currentStartedAt;
    const label = this.currentLabel;
    this.stopTimer();
    this.clearLine();
    const message = error instanceof Error ? error.message : String(error);
    process.stdout.write(`  ${color("31", "✗")} ${label} ${dim(formatMs(ms))}\n`);
    process.stdout.write(`    ${color("31", message)}\n`);
    this.currentLabel = null;
  }

  private render(): void {
    if (!this.isTty || this.currentLabel === null) return;
    const spinner = cyan(SPINNER_FRAMES[this.frame % SPINNER_FRAMES.length]);
    this.frame += 1;
    const elapsed = formatMs(now() - this.currentStartedAt);
    const done = this.finished.length;
    const fraction = done / this.totalSteps;
    const progress = dim(`${bar(fraction)} ${done}/${this.totalSteps}`);
    this.clearLine();
    process.stdout.write(`  ${spinner} ${this.currentLabel} ${dim(elapsed)}  ${progress}`);
  }

  private clearLine(): void {
    if (this.isTty) process.stdout.write("\r\x1b[K");
  }

  private stopTimer(): void {
    if (this.timer) {
      clearInterval(this.timer);
      this.timer = null;
    }
  }

  /** Print the closing summary with the total time and per-step breakdown. */
  finish(summary: { output: string; slides: number; warnings: number }): void {
    this.stopTimer();
    this.clearLine();
    const total = now() - this.startedAt;

    const slowest = [...this.finished].sort((a, b) => b.ms - a.ms).slice(0, 3);
    process.stdout.write(`\n${green("✓ Deck built")} ${dim("in " + formatMs(total))}\n`);
    process.stdout.write(`  ${dim("Output:")}   ${summary.output}\n`);
    process.stdout.write(`  ${dim("Slides:")}   ${summary.slides}\n`);
    process.stdout.write(
      `  ${dim("Warnings:")} ${summary.warnings === 0 ? dim("0") : color("33", String(summary.warnings))}\n`
    );
    if (slowest.length > 0) {
      process.stdout.write(`  ${dim("Slowest:")}  ${slowest.map((s) => `${s.label} ${formatMs(s.ms)}`).join(dim(", "))}\n`);
    }
    process.stdout.write("\n");
  }
}
