import { useEffect, useRef, useState } from "react";
import DownloadSection from "./DownloadSection";
import ProgressSection from "./ProgressSection";
import UploadSection from "./UploadSection";

const API_BASE = (import.meta.env.VITE_API_BASE || "").replace(/\/$/, "");

const INITIAL_STATUS = {
  job_id: null,
  status: "idle",
  progress_pct: 0,
  colleges_found: 0,
  branches_found: 0,
  rows_found: 0,
  current_college: "",
  current_branch: "",
  message: "",
  estimated_time_remaining_seconds: null,
  output_filename: null,
  partial_available: false,
  error: null,
};

export default function App() {
  const [selectedFile, setSelectedFile] = useState(null);
  const [dragActive, setDragActive] = useState(false);
  const [phase, setPhase] = useState("upload");
  const [status, setStatus] = useState(INITIAL_STATUS);
  const [error, setError] = useState("");
  const pollingRef = useRef(null);

  useEffect(() => {
    if (phase !== "processing" || !status.job_id) {
      return undefined;
    }

    const fetchStatus = async () => {
      try {
        const response = await fetch(`${API_BASE}/api/status/${status.job_id}`);
        const payload = await response.json();

        if (!response.ok) {
          throw new Error(payload.detail || "Failed to fetch job status.");
        }

        setStatus(payload);

        if (payload.status === "complete") {
          setPhase("complete");
        } else if (payload.status === "error") {
          setPhase("error");
        }
      } catch (requestError) {
        setError(requestError.message || "Unable to poll job status.");
        setPhase("upload");
      }
    };

    fetchStatus();
    pollingRef.current = window.setInterval(fetchStatus, 1500);

    return () => {
      if (pollingRef.current) {
        window.clearInterval(pollingRef.current);
      }
    };
  }, [phase, status.job_id]);

  const handleBrowse = (event) => {
    const file = event.target.files?.[0];

    if (!file) {
      return;
    }

    setError("");
    setSelectedFile(file);
  };

  const handleDrop = (event) => {
    event.preventDefault();
    setDragActive(false);

    const file = event.dataTransfer.files?.[0];

    if (!file) {
      return;
    }

    if (!file.name.toLowerCase().endsWith(".pdf")) {
      setError("Please upload a valid MHT-CET CAP Round PDF");
      return;
    }

    setError("");
    setSelectedFile(file);
  };

  const handleConvert = async () => {
    if (!selectedFile) {
      return;
    }

    const formData = new FormData();
    formData.append("file", selectedFile);

    setError("");
    setPhase("processing");
    setStatus({
      ...INITIAL_STATUS,
      message: "Uploading PDF...",
    });

    try {
      const response = await fetch(`${API_BASE}/api/upload`, {
        method: "POST",
        body: formData,
      });

      // Prevent crashes when backend returns HTML instead of JSON
      const contentType = response.headers.get("content-type") || "";

      if (!contentType.includes("application/json")) {
        throw new Error(
          `The server returned an unexpected response (status ${response.status}). ` +
            `The Space may be starting up or overloaded — wait 30 seconds and try again.`
        );
      }

      const payload = await response.json();

      if (!response.ok) {
        throw new Error(
          payload.detail || "Unable to upload the selected PDF."
        );
      }

      setStatus((currentStatus) => ({
        ...currentStatus,
        job_id: payload.job_id,
        status: payload.status,
        message: payload.message || "Reading PDF pages...",
      }));
    } catch (requestError) {
      setPhase("upload");
      setError(
        requestError.message || "Unable to upload the selected PDF."
      );
    }
  };

  const handleReset = () => {
    setSelectedFile(null);
    setDragActive(false);
    setPhase("upload");
    setStatus(INITIAL_STATUS);
    setError("");
  };

  return (
    <main className="min-h-screen px-4 py-10 text-ink sm:px-6 lg:px-10">
      <div className="mx-auto max-w-6xl">
        <header className="animate-rise rounded-[32px] bg-ink px-8 py-10 text-white shadow-panel">
          <div className="flex flex-col gap-8 lg:flex-row lg:items-end lg:justify-between">
            <div>
              <p className="text-sm font-semibold uppercase tracking-[0.28em] text-gold">
                Maharashtra Engineering Admissions
              </p>

              <h1 className="mt-4 max-w-3xl font-display text-4xl font-bold leading-tight sm:text-5xl">
                MHT-CET Cutoff PDF to Counseling Excel Converter
              </h1>

              <p className="mt-4 max-w-2xl text-sm leading-7 text-white/80 sm:text-base">
                Designed for educational counselors who need a ranked,
                filter-ready cutoff workbook in under a minute.
              </p>
            </div>

            <div className="grid gap-4 rounded-[28px] bg-white/10 p-5 backdrop-blur sm:grid-cols-2">
              <div>
                <p className="text-xs uppercase tracking-[0.18em] text-white/60">
                  Counseling promise
                </p>
                <p className="mt-2 text-xl font-bold">
                  10-second answer flow
                </p>
              </div>

              <div>
                <p className="text-xs uppercase tracking-[0.18em] text-white/60">
                  Output
                </p>
                <p className="mt-2 text-xl font-bold">
                  Excel for CAP filters
                </p>
              </div>
            </div>
          </div>
        </header>

        <div className="mt-8 grid gap-8 lg:grid-cols-[1.3fr,0.7fr]">
          <div className="space-y-8">
            {phase === "upload" ? (
              <UploadSection
                file={selectedFile}
                dragActive={dragActive}
                disabled={false}
                error={error}
                onBrowse={handleBrowse}
                onDrop={handleDrop}
                onDragStateChange={setDragActive}
                onConvert={handleConvert}
              />
            ) : null}

            {phase === "processing" ? (
              <ProgressSection status={status} />
            ) : null}

            {phase === "complete" || phase === "error" ? (
              <DownloadSection
                status={status}
                apiBase={API_BASE}
                onReset={handleReset}
              />
            ) : null}
          </div>

          <aside className="animate-rise rounded-[28px] border border-white/70 bg-white/75 p-6 shadow-panel backdrop-blur">
            <p className="text-sm font-semibold uppercase tracking-[0.24em] text-coral">
              Counselor workflow
            </p>

            <ol className="mt-4 space-y-4 text-sm leading-6 text-slate">
              <li className="rounded-2xl bg-mist p-4">
                <span className="block font-semibold text-ink">
                  1. Gather student context
                </span>
                Percentile, category, city or district, preferred branch,
                and gender.
              </li>

              <li className="rounded-2xl bg-mist p-4">
                <span className="block font-semibold text-ink">
                  2. Filter the workbook
                </span>
                Use city, district, branch, and college type filters to
                narrow options instantly.
              </li>

              <li className="rounded-2xl bg-mist p-4">
                <span className="block font-semibold text-ink">
                  3. Compare cutoffs
                </span>
                Sort by the relevant percentile column and identify safe and
                borderline colleges.
              </li>
            </ol>

            <div className="mt-6 rounded-[24px] bg-gradient-to-br from-gold/25 to-coral/10 p-5">
              <p className="text-xs uppercase tracking-[0.18em] text-slate">
                Current app state
              </p>

              <p className="mt-2 text-xl font-bold text-ink">
                {phase === "upload"
                  ? "Ready for a new PDF"
                  : phase === "processing"
                  ? "Conversion in progress"
                  : "Workbook ready"}
              </p>

              <p className="mt-2 text-sm leading-6 text-slate">
                {phase === "processing"
                  ? status.message || "Reading PDF pages..."
                  : "Upload the official cutoff PDF to generate a counselor-ready workbook."}
              </p>
            </div>
          </aside>
        </div>
      </div>
    </main>
  );
}
