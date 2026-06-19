"use client";

import { useState, useEffect } from "react";
import { format } from "date-fns";
import { motion } from "framer-motion";
import { Calendar, Clock, LayoutTemplate, Plus, Trash2, ArrowRight } from "lucide-react";
import toast from "react-hot-toast";
import Link from "next/link";
import { useRouter } from "next/navigation";

interface Project {
  id: string;
  title: string;
  start_date: string;
  calendar_format: string;
  created_at: string;
  activities: any[];
}

export default function GlobalDashboard() {
  const [projects, setProjects] = useState<Project[]>([]);
  const [loading, setLoading] = useState(true);
  const router = useRouter();

  useEffect(() => {
    fetchProjects();
  }, []);

  const fetchProjects = async () => {
    try {
      // Artificial delay to make skeleton loading visible for better UX
      await new Promise(r => setTimeout(r, 600));
      const res = await fetch("/api/projects");
      if (res.ok) {
        const data = await res.json();
        setProjects(data);
      }
    } catch (error) {
      console.error("Failed to fetch projects", error);
    } finally {
      setLoading(false);
    }
  };

  const createProject = async () => {
    const promise = fetch("/api/projects", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        title: "New Project",
        start_date: new Date().toISOString().split("T")[0],
        calendar_format: "5-day week"
      })
    }).then(async res => {
      if (!res.ok) throw new Error("Failed to create");
      const newProject = await res.json();
      router.push(`/projects/${newProject.id}`);
      return newProject;
    });

    toast.promise(promise, {
      loading: 'Creating project...',
      success: 'Project created!',
      error: 'Failed to create project.'
    });
  };

  const deleteProject = async (id: string, e: React.MouseEvent) => {
    e.preventDefault();
    toast((t) => (
      <div className="flex flex-col gap-3">
        <p className="font-medium text-slate-800">Are you sure you want to delete this project?</p>
        <div className="flex justify-end gap-2">
          <button onClick={() => toast.dismiss(t.id)} className="px-3 py-1.5 text-sm font-medium text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-md transition-colors">Cancel</button>
          <button onClick={() => { 
            toast.dismiss(t.id); 
            const promise = fetch(`/api/projects/${id}`, { method: "DELETE" }).then(res => {
              if(!res.ok) throw new Error();
              fetchProjects();
            });
            toast.promise(promise, { loading: 'Deleting...', success: 'Project deleted', error: 'Failed to delete' });
          }} className="px-3 py-1.5 text-sm font-medium text-white bg-red-600 hover:bg-red-700 rounded-md transition-colors">Delete</button>
        </div>
      </div>
    ), { duration: 5000 });
  };

  return (
    <div className="min-h-screen bg-slate-50 dark:bg-slate-950 p-8 text-slate-900 dark:text-slate-100">
      <div className="max-w-6xl mx-auto space-y-8">
        <header className="flex flex-col sm:flex-row sm:items-center justify-between gap-6">
          <div>
            <div className="flex items-center gap-4 mb-2">
              <img src="/Full-IESL-Logo.png" alt="IESL Logo" className="h-10 object-contain" />
            </div>
            <h1 className="text-3xl font-bold tracking-tight text-[#006634] dark:text-blue-400">
              Project Scheduler
            </h1>
            <p className="text-slate-500 dark:text-slate-400 mt-1">Manage your construction schedules and Gantt charts.</p>
          </div>
          <button
            onClick={createProject}
            className="flex items-center gap-2 bg-[#006634] hover:bg-[#004d26] text-white px-6 py-3 rounded-xl shadow-lg shadow-[#006634]/20 transition-all active:scale-95 font-medium"
          >
            <Plus size={20} />
            New Project
          </button>
        </header>

        {loading ? (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {[1, 2, 3].map(i => (
              <div key={i} className="h-64 bg-slate-200 dark:bg-slate-800 rounded-2xl animate-pulse"></div>
            ))}
          </div>
        ) : projects.length === 0 ? (
          <div className="flex flex-col items-center justify-center h-96 border-2 border-dashed border-slate-300 dark:border-slate-800 rounded-3xl bg-slate-100/50 dark:bg-slate-900/50">
            <LayoutTemplate size={48} className="text-slate-400 mb-4" />
            <h3 className="text-xl font-semibold mb-2">No projects yet</h3>
            <p className="text-slate-500 mb-6 text-center max-w-sm">Create your first project to start planning activities and generating Gantt charts.</p>
            <button onClick={createProject} className="text-[#006634] font-medium hover:underline">Create a project →</button>
          </div>
        ) : (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {projects.map((project, idx) => (
              <motion.div
                key={project.id}
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: idx * 0.1 }}
              >
                <Link href={`/projects/${project.id}`}>
                  <div className="group relative bg-white dark:bg-slate-900 rounded-2xl p-6 shadow-sm border border-slate-200 dark:border-slate-800 hover:shadow-xl hover:border-blue-500/50 transition-all duration-300 h-full flex flex-col cursor-pointer">
                    <button
                      onClick={(e) => deleteProject(project.id, e)}
                      className="absolute top-4 right-4 p-2 text-slate-400 hover:text-red-500 hover:bg-red-50 dark:hover:bg-red-500/10 rounded-lg transition-colors opacity-0 group-hover:opacity-100"
                    >
                      <Trash2 size={18} />
                    </button>
                    
                    <h3 className="text-xl font-bold mb-4 pr-8 line-clamp-2">{project.title}</h3>
                    
                    <div className="space-y-3 mt-auto">
                      <div className="flex items-center gap-3 text-sm text-slate-600 dark:text-slate-400">
                        <Calendar size={16} className="text-blue-500" />
                        <span>{project.start_date ? format(new Date(project.start_date), "MMM d, yyyy") : "Not set"}</span>
                      </div>
                      <div className="flex items-center gap-3 text-sm text-slate-600 dark:text-slate-400">
                        <Clock size={16} className="text-indigo-500" />
                        <span>{project.calendar_format}</span>
                      </div>
                      <div className="flex items-center gap-3 text-sm text-slate-600 dark:text-slate-400">
                        <LayoutTemplate size={16} className="text-purple-500" />
                        <span>{project.activities?.length || 0} activities</span>
                      </div>
                    </div>
                    
                    <div className="mt-6 pt-4 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between text-sm font-medium text-[#006634] dark:text-blue-400 opacity-0 group-hover:opacity-100 transition-opacity">
                      <span>Open Workspace</span>
                      <ArrowRight size={16} />
                    </div>
                  </div>
                </Link>
              </motion.div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
