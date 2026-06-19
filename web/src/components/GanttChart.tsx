"use client";

import { useEffect, useState, useRef } from "react";
import { Loader2, Image as ImageIcon } from "lucide-react";
import html2canvas from "html2canvas";
import { format, parseISO } from "date-fns";

interface GanttData {
  project_start_date: string;
  timeline_start: string;
  timeline_days: number;
  date_timeline: string[];
  tasks: any[];
}

export default function GanttChart({ projectId }: { projectId: string }) {
  const [data, setData] = useState<GanttData | null>(null);
  const [loading, setLoading] = useState(true);
  const [exporting, setExporting] = useState(false);
  const [viewScale, setViewScale] = useState<'day' | 'week' | 'month'>('day');
  const chartRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    fetchGanttData();
  }, [projectId]);

  const fetchGanttData = async () => {
    try {
      // Artificial delay so the loading spinner is visible for a moment
      await new Promise(r => setTimeout(r, 600));
      const res = await fetch(`/api/projects/${projectId}/gantt`);
      if (res.ok) {
        setData(await res.json());
      }
    } catch (error) {
      console.error("Failed to fetch Gantt data", error);
    } finally {
      setLoading(false);
    }
  };

  const exportAsImage = async () => {
    if (!chartRef.current) return;
    setExporting(true);
    
    const parent = chartRef.current.parentElement;
    const parentOriginalOverflow = parent ? parent.style.overflow : '';
    
    // Find all sticky and truncate elements to temporarily fix them for html2canvas
    const stickyElements = chartRef.current.querySelectorAll('.sticky');
    const truncateElements = chartRef.current.querySelectorAll('.truncate');
    
    try {
      if (parent) {
        parent.style.overflow = 'visible';
      }
      
      stickyElements.forEach((el) => {
        (el as HTMLElement).style.setProperty('position', 'relative', 'important');
      });
      
      truncateElements.forEach((el) => {
        (el as HTMLElement).classList.remove('truncate');
        (el as HTMLElement).style.whiteSpace = 'nowrap';
        (el as HTMLElement).style.overflow = 'visible';
      });

      // Small delay to let browser reflow
      await new Promise(resolve => setTimeout(resolve, 100));

      const canvas = await html2canvas(chartRef.current, {
        scale: 2,
        backgroundColor: "#ffffff",
        logging: false,
        width: chartRef.current.scrollWidth,
        height: chartRef.current.scrollHeight
      });
      const url = canvas.toDataURL("image/png");
      const a = document.createElement("a");
      a.href = url;
      a.download = `Gantt_Chart_${projectId}.png`;
      document.body.appendChild(a);
      a.click();
      a.remove();
    } catch (err) {
      console.error("Export failed", err);
    } finally {
      if (parent) parent.style.overflow = parentOriginalOverflow;
      stickyElements.forEach((el) => {
        (el as HTMLElement).style.removeProperty('position');
      });
      truncateElements.forEach((el) => {
        (el as HTMLElement).classList.add('truncate');
        (el as HTMLElement).style.removeProperty('white-space');
        (el as HTMLElement).style.removeProperty('overflow');
      });
      setExporting(false);
    }
  };

  if (loading) {
    return <div className="flex-1 flex items-center justify-center min-h-[400px]"><Loader2 className="animate-spin text-[#006634]" size={32} /></div>;
  }

  if (!data || data.tasks.length === 0) {
    return <div className="flex-1 flex items-center justify-center min-h-[400px] text-slate-500">No activities found to generate chart.</div>;
  }

  // Calculate grid layout constants
  const CELL_WIDTH = viewScale === 'day' ? 24 : viewScale === 'week' ? 8 : 2;
  const timelineWidth = data.timeline_days * CELL_WIDTH;

  return (
    <div className="flex flex-col h-full overflow-hidden relative border border-slate-200 dark:border-slate-800 rounded-xl bg-white dark:bg-slate-900 shadow-sm m-6">
      <div className="flex justify-between items-center p-4 border-b border-slate-200 dark:border-slate-800 shrink-0 bg-slate-50 dark:bg-slate-800/50">
        <div className="flex items-center gap-4">
            <h3 className="font-semibold text-lg">Timeline View</h3>
            <div className="flex bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-lg p-1 ml-4 shadow-sm">
              {['day', 'week', 'month'].map(s => (
                <button key={s} onClick={() => setViewScale(s as any)} className={`px-3 py-1 text-xs font-medium rounded-md capitalize transition-colors ${viewScale === s ? 'bg-slate-100 dark:bg-slate-800 text-slate-900 dark:text-white' : 'text-slate-500 hover:text-slate-700 dark:hover:text-slate-300'}`}>
                  {s}
                </button>
              ))}
            </div>
            <div className="flex gap-4 text-xs font-medium text-slate-500 ml-4">
                <span className="flex items-center gap-1"><div className="w-3 h-3 rounded-sm bg-[#5B9BD5]"></div> Goal</span>
                <span className="flex items-center gap-1"><div className="w-3 h-3 rounded-sm bg-[#EAB308]"></div> Milestone</span>
                <span className="flex items-center gap-1"><div className="w-3 h-3 rounded-sm bg-red-500"></div> Critical Path</span>
            </div>
        </div>
        <button
          onClick={exportAsImage}
          disabled={exporting}
          className="flex items-center gap-2 bg-[#006634]/10 hover:bg-[#006634]/20 text-[#006634] px-4 py-2 rounded-lg transition-colors font-medium text-sm disabled:opacity-50"
        >
          {exporting ? <Loader2 size={16} className="animate-spin" /> : <ImageIcon size={16} />}
          {exporting ? "Exporting..." : "Save Image"}
        </button>
      </div>

      <div className="flex-1 overflow-auto custom-scrollbar bg-white dark:bg-slate-900 rounded-b-xl">
        <div 
          ref={chartRef} 
          className="min-w-max pb-6 pr-6"
        >
          {/* Header Row: Months/Dates */}
          <div className="flex sticky top-0 z-20 bg-white pb-2 border-b border-slate-200 pt-6">
            <div className="w-64 shrink-0 font-medium text-xs text-slate-500 uppercase tracking-wider pl-6 sticky left-0 z-30 bg-white border-r border-slate-100">
              Activity Name
            </div>
            <div className="flex relative" style={{ width: timelineWidth }}>
              {data.date_timeline.map((dateStr, i) => {
                const date = parseISO(dateStr);
                const isFirstOfMonth = date.getDate() === 1 || i === 0;
                const isMonday = date.getDay() === 1;
                return (
                  <div key={i} className={`flex flex-col items-center shrink-0 ${viewScale === 'day' ? 'border-l border-slate-200' : ''}`} style={{ width: CELL_WIDTH }}>
                    <div className="h-6 text-[10px] text-slate-400 font-semibold w-full text-center relative">
                      {isFirstOfMonth ? (
                        <span className="absolute -left-2 bg-white px-1 z-10 text-indigo-600">
                          {format(date, "MMM")} {viewScale === 'month' ? format(date, "yyyy") : ""}
                        </span>
                      ) : null}
                    </div>
                    {viewScale === 'day' && <div className="text-[10px] text-slate-500">{format(date, "d")}</div>}
                    {viewScale === 'week' && isMonday && <div className="text-[10px] text-slate-500 relative -left-3">W{format(date, "w")}</div>}
                  </div>
                );
              })}
            </div>
          </div>

          {/* Task Rows */}
          <div className="pt-2 bg-white relative z-0">
            {/* Watermark Overlay (Single Centered) */}
            <div 
              className="absolute top-2 bottom-0 pointer-events-none z-0 opacity-[0.06]"
              style={{
                left: '256px',
                width: timelineWidth,
                backgroundImage: "url('/Half-IESL-Logo.png')",
                backgroundRepeat: "no-repeat",
                backgroundSize: "50%",
                backgroundPosition: "center center",
                filter: "grayscale(100%)"
              }}
            />
            {/* Background Grid rendered ONCE for entire chart to save DOM elements */}
            <div className="absolute top-2 bottom-0 flex pointer-events-none opacity-20 z-0" style={{ left: '256px', width: timelineWidth }}>
              {data.date_timeline.map((dateStr, i) => {
                const date = parseISO(dateStr);
                const showBorder = viewScale === 'day' || (viewScale === 'week' && date.getDay() === 1) || (viewScale === 'month' && date.getDate() === 1);
                return (
                  <div key={i} className={`${showBorder ? 'border-l border-slate-300' : ''} h-full border-solid`} style={{ width: CELL_WIDTH }}></div>
                );
              })}
            </div>

            <div className="relative z-10">
              {data.tasks.map((task, idx) => {
                // Calculate offset in days
                const timelineStartMs = parseISO(data.timeline_start).getTime();
                const taskStartMs = parseISO(task.start_date).getTime();
                const offsetDays = Math.round((taskStartMs - timelineStartMs) / 86400000);
                
                const leftPos = offsetDays * CELL_WIDTH;
                const barWidth = Math.max((task.end_day - task.start_day) * CELL_WIDTH, 4);

                let bgColor = "bg-[#EAB308] text-white"; // Solid visible yellow
                if (task.is_critical) {
                  bgColor = task.task_type === "Goal" ? "bg-red-500" : "bg-red-700";
                } else if (task.task_type === "Goal") {
                  bgColor = "bg-[#5B9BD5]";
                }

                return (
                  <div key={idx} className="flex items-center py-2 border-b border-slate-100 hover:bg-slate-50 transition-colors group">
                    <div className="w-64 shrink-0 text-sm font-medium pr-4 pl-6 truncate text-slate-700 group-hover:text-slate-900 transition-colors sticky left-0 z-10 bg-white border-r border-slate-100" title={task.name}>
                      {idx + 1}. {task.name}
                    </div>
                    <div className="relative h-8 flex items-center" style={{ width: timelineWidth }}>
                      <div 
                        className={`absolute h-5 rounded-sm ${bgColor} shadow-sm flex items-center justify-center text-[10px] text-white font-bold overflow-hidden cursor-pointer hover:brightness-110 transition-all`}
                        style={{ left: leftPos, width: barWidth }}
                        title={`${task.name} (${task.duration} days)`}
                      >
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
