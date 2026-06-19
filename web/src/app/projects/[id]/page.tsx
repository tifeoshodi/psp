"use client";

import { useState, useEffect } from "react";
import { useParams, useRouter } from "next/navigation";
import { ArrowLeft, Save, Plus, Trash2, Download, Eye, FileSpreadsheet, Loader2, LayoutList, CalendarRange, PanelLeftClose, PanelLeftOpen } from "lucide-react";
import Link from "next/link";
import { motion, AnimatePresence } from "framer-motion";
import toast from "react-hot-toast";
import GanttChart from "@/components/GanttChart";

export default function ProjectWorkspace() {
  const { id } = useParams() as { id: string };
  const router = useRouter();

  const [project, setProject] = useState<any>(null);
  const [activities, setActivities] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  
  // Modal states
  const [previewOpen, setPreviewOpen] = useState(false);
  const [generating, setGenerating] = useState(false);
  
  // View mode
  const [viewMode, setViewMode] = useState<'list' | 'gantt'>('list');
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);

  // New activity form
  const [newActivity, setNewActivity] = useState({
    task: "",
    action_needed: "",
    duration: 1,
    precursor: "",
    sequence: 1,
    resources: "",
    budget: 0,
    section: "Post Kick-off Activities"
  });

  useEffect(() => {
    fetchProject();
  }, [id]);

  const fetchProject = async () => {
    try {
      // Artificial delay so the loading spinner is visible for a moment
      await new Promise(r => setTimeout(r, 600));
      const res = await fetch(`/api/projects/${id}`);
      if (res.ok) {
        const data = await res.json();
        setProject(data);
        setActivities(data.activities || []);
      } else {
        router.push("/");
      }
    } catch (error) {
      console.error(error);
    } finally {
      setLoading(false);
    }
  };

  const saveProject = async () => {
    setSaving(true);
    const promise = fetch(`/api/projects/${id}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        title: project.title,
        start_date: project.start_date,
        calendar_format: project.calendar_format,
        logo_path: project.logo_path
      })
    }).then(res => {
      if (!res.ok) throw new Error();
    }).finally(() => setSaving(false));

    toast.promise(promise, {
      loading: 'Saving...',
      success: 'Project saved!',
      error: 'Failed to save project.'
    });
  };

  const addActivity = async (e: React.FormEvent) => {
    e.preventDefault();
    const promise = fetch(`/api/projects/${id}/activities`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
          ...newActivity,
          duration: Number(newActivity.duration),
          sequence: Number(newActivity.sequence),
          budget: Number(newActivity.budget)
      })
    }).then(async res => {
      if (!res.ok) throw new Error();
      const added = await res.json();
      setActivities([...activities, added]);
      setNewActivity({
        ...newActivity,
        task: "",
        action_needed: "",
        sequence: Number(newActivity.sequence) + 1
      });
    });

    toast.promise(promise, {
      loading: 'Adding activity...',
      success: 'Activity added!',
      error: 'Failed to add activity.'
    });
  };

  const confirmDeleteActivity = (activityId: string) => {
    toast((t) => (
      <div className="flex flex-col gap-3">
        <p className="font-medium text-slate-800">Are you sure you want to delete this activity?</p>
        <div className="flex justify-end gap-2">
          <button onClick={() => toast.dismiss(t.id)} className="px-3 py-1.5 text-sm font-medium text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-md transition-colors">Cancel</button>
          <button onClick={() => { 
            toast.dismiss(t.id); 
            const promise = fetch(`/api/projects/${id}/activities/${activityId}`, { method: "DELETE" }).then(res => {
              if(!res.ok) throw new Error();
              setActivities(activities.filter(a => a.id !== activityId));
            });
            toast.promise(promise, { loading: 'Deleting...', success: 'Activity deleted', error: 'Failed to delete' });
          }} className="px-3 py-1.5 text-sm font-medium text-white bg-red-600 hover:bg-red-700 rounded-md transition-colors">Delete</button>
        </div>
      </div>
    ), { duration: 5000 });
  };

  const generateExcel = async () => {
    setGenerating(true);
    const promise = fetch(`/api/projects/${id}/generate-excel`, { method: "POST" }).then(async res => {
      if (!res.ok) throw new Error();
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${project.title.replace(/\s+/g, '_')}_Schedule.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    }).finally(() => {
      setGenerating(false);
      setPreviewOpen(false);
    });

    toast.promise(promise, {
      loading: 'Generating Excel...',
      success: 'Excel generated successfully!',
      error: 'Failed to generate Excel file.'
    });
  };

  if (loading) return <div className="min-h-screen flex items-center justify-center bg-slate-50 dark:bg-slate-950 dark:text-white"><Loader2 className="animate-spin text-[#006634]" size={48} /></div>;
  if (!project) return null;

  return (
    <div className="h-screen bg-slate-50 dark:bg-slate-950 text-slate-900 dark:text-slate-100 flex flex-col">
      {/* Top Bar */}
      <header className="sticky top-0 z-10 bg-white/80 dark:bg-slate-900/80 backdrop-blur-md border-b border-slate-200 dark:border-slate-800 px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-4">
          <Link href="/" className="p-2 hover:bg-slate-100 dark:hover:bg-slate-800 rounded-full transition-colors text-slate-500">
            <ArrowLeft size={20} />
          </Link>
          <img src="/Full-IESL-Logo.png" alt="IESL Logo" className="h-8 object-contain hidden sm:block" />
          <input
            type="text"
            value={project.title}
            onChange={e => setProject({ ...project, title: e.target.value })}
            className="text-2xl font-bold bg-transparent border-none focus:ring-0 p-0 hover:bg-slate-100 dark:hover:bg-slate-800 rounded px-2 transition-colors cursor-text w-96 text-[#006634] dark:text-blue-400 ml-2"
          />
        </div>
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2">
            <input
              type="date"
              value={project.start_date || ""}
              onChange={e => setProject({ ...project, start_date: e.target.value })}
              className="bg-slate-100 dark:bg-slate-800 border-none rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500"
            />
            <select
              value={project.calendar_format}
              onChange={e => setProject({ ...project, calendar_format: e.target.value })}
              className="bg-slate-100 dark:bg-slate-800 border-none rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500"
            >
              <option value="5-day week">5-day week</option>
              <option value="6-day week">6-day week</option>
              <option value="7-day week">7-day week</option>
            </select>
          </div>
          
          <div className="flex bg-slate-100 dark:bg-slate-800 p-1 rounded-lg border border-slate-200 dark:border-slate-700">
            <button 
              onClick={() => setViewMode('list')}
              className={`flex items-center gap-2 px-3 py-1.5 text-sm font-medium rounded-md transition-colors ${viewMode === 'list' ? 'bg-white dark:bg-slate-700 shadow-sm text-slate-900 dark:text-white' : 'text-slate-500 hover:text-slate-700 dark:hover:text-slate-300'}`}
            >
              <LayoutList size={16} /> List
            </button>
            <button 
              onClick={() => setViewMode('gantt')}
              className={`flex items-center gap-2 px-3 py-1.5 text-sm font-medium rounded-md transition-colors ${viewMode === 'gantt' ? 'bg-white dark:bg-slate-700 shadow-sm text-slate-900 dark:text-white' : 'text-slate-500 hover:text-slate-700 dark:hover:text-slate-300'}`}
            >
              <CalendarRange size={16} /> Gantt
            </button>
          </div>
          
          <button onClick={saveProject} className="p-2 text-[#006634] hover:bg-[#006634]/10 dark:hover:bg-[#006634]/20 rounded-lg transition-colors">
            {saving ? <Loader2 className="animate-spin" size={20} /> : <Save size={20} />}
          </button>
          <button onClick={() => setPreviewOpen(true)} className="flex items-center gap-2 bg-[#006634] hover:bg-[#004d26] text-white px-4 py-2 rounded-lg transition-colors shadow-lg shadow-[#006634]/20 text-sm font-medium">
            <FileSpreadsheet size={18} />
            Export
          </button>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 flex overflow-hidden">
        {/* Sidebar Form */}
        <motion.aside 
          initial={false}
          animate={{ width: isSidebarOpen ? 320 : 64 }}
          transition={{ duration: 0.25, ease: "easeInOut" }}
          className="border-r border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 hidden md:flex flex-col shrink-0 overflow-hidden"
        >
          {/* Header inside Sidebar */}
          <div className="flex items-center h-16 px-3 shrink-0">
            <AnimatePresence>
              {isSidebarOpen && (
                <motion.div 
                  initial={{ opacity: 0, width: 0 }}
                  animate={{ opacity: 1, width: "auto" }}
                  exit={{ opacity: 0, width: 0 }}
                  transition={{ duration: 0.2 }}
                  className="flex-1 overflow-hidden whitespace-nowrap"
                >
                  <h2 className="font-semibold text-lg flex items-center gap-2 pl-2 text-slate-800 dark:text-slate-100">
                    <Plus size={18} className="text-[#006634] shrink-0" />
                    Add Activity
                  </h2>
                </motion.div>
              )}
            </AnimatePresence>
            <div className={`w-10 h-10 flex items-center justify-center shrink-0 ${isSidebarOpen ? 'ml-auto' : 'mx-auto'}`}>
              <button 
                onClick={() => setIsSidebarOpen(!isSidebarOpen)}
                className="p-2 hover:bg-slate-100 dark:hover:bg-slate-800 rounded-lg text-slate-500 transition-colors"
                title={isSidebarOpen ? "Collapse Panel" : "Expand Panel"}
              >
                {isSidebarOpen ? <PanelLeftClose size={20} /> : <PanelLeftOpen size={20} />}
              </button>
            </div>
          </div>

          <div className="flex-1 overflow-y-auto custom-scrollbar overflow-x-hidden">
            <div 
              className="w-80 p-5 pt-0 transition-opacity duration-200" 
              style={{ opacity: isSidebarOpen ? 1 : 0, pointerEvents: isSidebarOpen ? 'auto' : 'none' }}
            >
              <form onSubmit={addActivity} className="space-y-3">
                <div>
                  <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Task Name</label>
                  <input required type="text" value={newActivity.task} onChange={e => setNewActivity({...newActivity, task: e.target.value})} className="w-full bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 outline-none" />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Action Needed</label>
                  <input type="text" value={newActivity.action_needed} onChange={e => setNewActivity({...newActivity, action_needed: e.target.value})} className="w-full bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 outline-none" />
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Duration (days)</label>
                    <input required type="number" min="0" value={newActivity.duration} onChange={e => setNewActivity({...newActivity, duration: e.target.value as any})} className="w-full bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 outline-none" />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Sequence</label>
                    <input required type="number" min="1" value={newActivity.sequence} onChange={e => setNewActivity({...newActivity, sequence: e.target.value as any})} className="w-full bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 outline-none" />
                  </div>
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Precursor</label>
                  <input type="text" placeholder="Activity ID or Name" value={newActivity.precursor} onChange={e => setNewActivity({...newActivity, precursor: e.target.value})} className="w-full bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 outline-none" />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Section</label>
                  <select value={newActivity.section} onChange={e => setNewActivity({...newActivity, section: e.target.value})} className="w-full bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 outline-none">
                    <option value="Pre-Kickoff Activities">Pre-Kickoff</option>
                    <option value="Post Kick-off Activities">Post Kick-off</option>
                  </select>
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Budget</label>
                    <input type="number" min="0" step="0.01" value={newActivity.budget} onChange={e => setNewActivity({...newActivity, budget: e.target.value as any})} className="w-full bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 outline-none" />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Resources</label>
                    <input type="text" value={newActivity.resources} onChange={e => setNewActivity({...newActivity, resources: e.target.value})} className="w-full bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 outline-none" />
                  </div>
                </div>
                <button type="submit" className="w-full bg-[#006634] hover:bg-[#004d26] text-white font-medium py-2 rounded-lg mt-4 transition-colors shadow-sm">
                  Add to Schedule
                </button>
              </form>
            </div>
          </div>
        </motion.aside>

        {/* Data Table / Gantt Chart */}
        <section className="flex-1 flex flex-col min-h-0 bg-slate-50/50 dark:bg-slate-950 p-6 overflow-hidden">
          {viewMode === 'list' ? (
            <div className="flex-1 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl overflow-hidden shadow-sm flex flex-col min-h-0">
              <div className="flex-1 overflow-auto custom-scrollbar">
              <table className="w-full text-sm text-left">
                <thead className="bg-slate-50 dark:bg-slate-800/50 text-slate-500 dark:text-slate-400 font-medium border-b border-slate-200 dark:border-slate-800">
                  <tr>
                    <th className="px-4 py-3">Seq</th>
                    <th className="px-4 py-3">Task</th>
                    <th className="px-4 py-3 text-center">Duration</th>
                    <th className="px-4 py-3">Precursor</th>
                    <th className="px-4 py-3">Budget</th>
                    <th className="px-4 py-3 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {(() => {
                    if (activities.length === 0) {
                      return (
                        <tr>
                          <td colSpan={6} className="px-4 py-12 text-center text-slate-500">No activities added yet.</td>
                        </tr>
                      );
                    }
                    
                    const sorted = [...activities].sort((a,b) => {
                      if (a.section !== b.section) {
                        return a.section.includes('Pre') ? -1 : 1;
                      }
                      return a.sequence - b.sequence;
                    });
                    
                    let currentSection = "";
                    return sorted.map((act) => {
                      const rows = [];
                      if (act.section !== currentSection) {
                        currentSection = act.section;
                        const isPreKickoff = act.section.toLowerCase().includes('pre');
                        const bgColorClass = isPreKickoff ? 'bg-[#ffff99] dark:bg-[#ffff99]/20' : 'bg-[#D2E3A3] dark:bg-[#D2E3A3]/20';
                        const borderColorClass = isPreKickoff ? 'border-[#e6e68a] dark:border-[#ffff99]/30' : 'border-[#bdcc93] dark:border-[#D2E3A3]/30';
                        
                        rows.push(
                          <tr key={`header-${act.section}`} className={`${bgColorClass} border-y ${borderColorClass}`}>
                            <td colSpan={6} className="px-4 py-2 font-semibold text-slate-900 dark:text-slate-100 text-xs uppercase tracking-wider">
                              {act.section}
                            </td>
                          </tr>
                        );
                      }
                      
                      rows.push(
                      <tr key={act.id} className="border-b border-slate-100 dark:border-slate-800/50 hover:bg-slate-50 dark:hover:bg-slate-800/20 transition-colors">
                        <td className="px-4 py-3 text-slate-500">#{act.sequence}</td>
                        <td className="px-4 py-3 font-medium">{act.task}</td>
                        <td className="px-4 py-3 text-center">{act.duration}d</td>
                        <td className="px-4 py-3 text-slate-500">{act.precursor || "-"}</td>
                        <td className="px-4 py-3">${Number(act.budget).toLocaleString()}</td>
                        <td className="px-4 py-3 text-right">
                          <button onClick={() => confirmDeleteActivity(act.id)} className="text-slate-400 hover:text-red-500 transition-colors p-1">
                            <Trash2 size={16} />
                          </button>
                        </td>
                      </tr>
                      );
                      
                      return rows;
                    });
                  })()}
                </tbody>
              </table>
              </div>
            </div>
          ) : (
            <div className="flex-1 min-h-0 -m-6">
                <GanttChart projectId={id} />
            </div>
          )}
        </section>
      </main>

      {/* Preview & Export Modal */}
      <AnimatePresence>
        {previewOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 sm:p-6">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
              onClick={() => !generating && setPreviewOpen(false)}
            />
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="relative bg-white dark:bg-slate-900 rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden border border-slate-200 dark:border-slate-800 flex flex-col max-h-[90vh]"
            >
              <div className="p-6 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between">
                <div>
                  <h3 className="text-xl font-bold flex items-center gap-2"><Eye className="text-indigo-500" /> Export Preview</h3>
                  <p className="text-sm text-slate-500 mt-1">Review details before generating the Excel file.</p>
                </div>
              </div>
              
              <div className="p-6 overflow-y-auto space-y-6 flex-1">
                <div className="grid grid-cols-2 gap-4">
                  <div className="bg-slate-50 dark:bg-slate-800/50 p-4 rounded-xl">
                    <p className="text-xs text-slate-500 uppercase font-semibold">Total Activities</p>
                    <p className="text-3xl font-bold mt-1 text-[#006634]">{activities.length}</p>
                  </div>
                  <div className="bg-slate-50 dark:bg-slate-800/50 p-4 rounded-xl">
                    <p className="text-xs text-slate-500 uppercase font-semibold">Total Budget</p>
                    <p className="text-3xl font-bold mt-1 text-emerald-600 dark:text-emerald-400">
                      ${activities.reduce((acc, a) => acc + Number(a.budget), 0).toLocaleString()}
                    </p>
                  </div>
                </div>

                <div className="space-y-2">
                   <h4 className="font-medium">Logo Upload (Optional)</h4>
                   <p className="text-sm text-slate-500">Currently using the default IESL logo. Logo upload feature requires cloud storage configuration.</p>
                </div>
              </div>

              <div className="p-6 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-900/50 flex justify-end gap-3">
                <button 
                  onClick={() => setPreviewOpen(false)} 
                  disabled={generating}
                  className="px-5 py-2.5 text-sm font-medium rounded-xl hover:bg-slate-200 dark:hover:bg-slate-800 transition-colors disabled:opacity-50"
                >
                  Cancel
                </button>
                <button
                  onClick={generateExcel}
                  disabled={generating || activities.length === 0}
                  className="px-5 py-2.5 text-sm font-medium rounded-xl bg-[#006634] hover:bg-[#004d26] text-white shadow-lg shadow-[#006634]/20 transition-all active:scale-95 disabled:opacity-50 flex items-center gap-2"
                >
                  {generating ? (
                    <><Loader2 size={18} className="animate-spin" /> Generating...</>
                  ) : (
                    <><Download size={18} /> Generate Excel & Gantt</>
                  )}
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
