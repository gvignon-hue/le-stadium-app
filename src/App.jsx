import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";
import {
  ChevronLeft, ChevronRight, X, Trash2, MessageSquare,
  Users, ChevronDown, Check, Plus, Pencil, LayoutGrid, Table2,
} from "lucide-react";

const LOGO_B64 = "data:image/png;base64,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";

const MANAGERS = ["Pierre Arnaud", "Greg"];
const DAY_NAMES = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"];
const DAY_SHORT = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"];

const RANGE_START = 8 * 60;
const RANGE_END = 23 * 60 + 30;
const RANGE_TOTAL = RANGE_END - RANGE_START;
const HOURS = Array.from({ length: 16 }, (_, i) => 8 + i);

const PRESETS = [
  { key: "matin", label: "Matin", start: 9 * 60 + 30, end: 16 * 60 + 30 },
  { key: "journee", label: "Journée", start: 11 * 60, end: 18 * 60 },
  { key: "soiree", label: "Soirée", start: 16 * 60 + 30, end: 23 * 60 + 30 },
  { key: "weekend", label: "Weekend", start: 9 * 60 + 30, end: 20 * 60 + 30 },
];

const EMP_PALETTE = [
  { bg: "bg-orange-500", text: "text-orange-700" },
  { bg: "bg-yellow-500", text: "text-yellow-700" },
  { bg: "bg-red-400", text: "text-red-700" },
  { bg: "bg-blue-600", text: "text-blue-700" },
  { bg: "bg-green-600", text: "text-green-700" },
  { bg: "bg-violet-600", text: "text-violet-700" },
  { bg: "bg-slate-500", text: "text-slate-700" },
  { bg: "bg-teal-600", text: "text-teal-700" },
  { bg: "bg-fuchsia-600", text: "text-fuchsia-700" },
  { bg: "bg-lime-600", text: "text-lime-700" },
];

const EMPLOYEES_SEED = [
  { id: "emp-1", name: "Aurélie", colorIdx: 0 },
  { id: "emp-2", name: "Nathan", colorIdx: 1 },
  { id: "emp-3", name: "Ewan", colorIdx: 2 },
  { id: "emp-4", name: "Tom", colorIdx: 3 },
  { id: "emp-5", name: "Blanche", colorIdx: 4 },
];

const PLANNING_SEED = {
  "2026-11-02": [
    { id: "seedp4-1", day: 0, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-2", day: 0, employee: "Nathan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-3", day: 1, employee: "Nathan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-4", day: 2, employee: "Nathan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-5", day: 3, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-6", day: 3, employee: "Ewan", start: 810, end: 1380, comment: "" },
    { id: "seedp4-7", day: 4, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-8", day: 5, employee: "Ewan", start: 570, end: 1230, comment: "" },
    { id: "seedp4-9", day: 6, employee: "Blanche", start: 570, end: 1230, comment: "" },
  ],
  "2026-11-09": [
    { id: "seedp4-10", day: 0, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-11", day: 0, employee: "Tom", start: 990, end: 1380, comment: "" },
    { id: "seedp4-12", day: 1, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-13", day: 1, employee: "Tom", start: 990, end: 1380, comment: "" },
    { id: "seedp4-14", day: 2, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-15", day: 2, employee: "Tom", start: 990, end: 1380, comment: "" },
    { id: "seedp4-16", day: 2, employee: "Ewan", start: 690, end: 1110, comment: "" },
    { id: "seedp4-17", day: 3, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-18", day: 3, employee: "Nathan", start: 810, end: 1380, comment: "" },
    { id: "seedp4-19", day: 3, employee: "Ewan", start: 690, end: 1110, comment: "" },
    { id: "seedp4-20", day: 4, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-21", day: 4, employee: "Ewan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-22", day: 5, employee: "Nathan", start: 570, end: 1230, comment: "" },
    { id: "seedp4-23", day: 6, employee: "Blanche", start: 570, end: 1230, comment: "" },
  ],
  "2026-11-16": [
    { id: "seedp4-24", day: 0, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-25", day: 0, employee: "Nathan", start: 1020, end: 1380, comment: "" },
    { id: "seedp4-26", day: 1, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-27", day: 1, employee: "Nathan", start: 1020, end: 1380, comment: "" },
    { id: "seedp4-28", day: 2, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-29", day: 2, employee: "Tom", start: 1080, end: 1380, comment: "" },
    { id: "seedp4-30", day: 2, employee: "Ewan", start: 570, end: 990, comment: "" },
    { id: "seedp4-31", day: 3, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-32", day: 3, employee: "Ewan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-33", day: 4, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-34", day: 4, employee: "Tom", start: 1080, end: 1380, comment: "" },
    { id: "seedp4-35", day: 4, employee: "Ewan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-36", day: 5, employee: "Nathan", start: 510, end: 990, comment: "" },
    { id: "seedp4-37", day: 5, employee: "Tom", start: 570, end: 1230, comment: "" },
    { id: "seedp4-38", day: 6, employee: "Blanche", start: 570, end: 1230, comment: "" },
  ],
  "2026-11-23": [
    { id: "seedp4-39", day: 0, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-40", day: 0, employee: "Tom", start: 990, end: 1380, comment: "" },
    { id: "seedp4-41", day: 1, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-42", day: 1, employee: "Tom", start: 990, end: 1380, comment: "" },
    { id: "seedp4-43", day: 2, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-44", day: 2, employee: "Nathan", start: 570, end: 990, comment: "" },
    { id: "seedp4-45", day: 2, employee: "Tom", start: 690, end: 1110, comment: "" },
    { id: "seedp4-46", day: 2, employee: "Ewan", start: 870, end: 1380, comment: "" },
    { id: "seedp4-47", day: 3, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-48", day: 3, employee: "Nathan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-49", day: 4, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-50", day: 4, employee: "Nathan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-51", day: 5, employee: "Ewan", start: 510, end: 1230, comment: "" },
    { id: "seedp4-52", day: 6, employee: "Blanche", start: 570, end: 1230, comment: "" },
  ],
  "2026-11-30": [
    { id: "seedp4-53", day: 0, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-54", day: 0, employee: "Tom", start: 990, end: 1380, comment: "" },
    { id: "seedp4-55", day: 1, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-56", day: 1, employee: "Tom", start: 900, end: 1380, comment: "" },
    { id: "seedp4-57", day: 2, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-58", day: 2, employee: "Nathan", start: 870, end: 1380, comment: "" },
    { id: "seedp4-59", day: 2, employee: "Ewan", start: 810, end: 1230, comment: "" },
    { id: "seedp4-60", day: 3, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-61", day: 3, employee: "Ewan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-62", day: 4, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-63", day: 4, employee: "Tom", start: 1080, end: 1380, comment: "" },
    { id: "seedp4-64", day: 4, employee: "Ewan", start: 690, end: 1110, comment: "" },
    { id: "seedp4-65", day: 5, employee: "Nathan", start: 510, end: 1230, comment: "" },
    { id: "seedp4-66", day: 6, employee: "Blanche", start: 570, end: 1230, comment: "" },
  ],
  "2026-12-07": [
    { id: "seedp4-67", day: 0, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-68", day: 0, employee: "Tom", start: 1080, end: 1380, comment: "" },
    { id: "seedp4-69", day: 1, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-70", day: 1, employee: "Tom", start: 1080, end: 1380, comment: "" },
    { id: "seedp4-71", day: 2, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-72", day: 2, employee: "Nathan", start: 570, end: 990, comment: "" },
    { id: "seedp4-73", day: 2, employee: "Ewan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-74", day: 3, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-75", day: 3, employee: "Nathan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-76", day: 3, employee: "Ewan", start: 690, end: 1110, comment: "" },
    { id: "seedp4-77", day: 4, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-78", day: 4, employee: "Nathan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-79", day: 4, employee: "Ewan", start: 570, end: 990, comment: "" },
    { id: "seedp4-80", day: 5, employee: "Tom", start: 510, end: 1230, comment: "" },
    { id: "seedp4-81", day: 6, employee: "Blanche", start: 570, end: 1230, comment: "" },
  ],
  "2026-12-14": [
    { id: "seedp4-82", day: 0, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-83", day: 0, employee: "Nathan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-84", day: 0, employee: "Tom", start: 570, end: 990, comment: "" },
    { id: "seedp4-85", day: 1, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-86", day: 1, employee: "Nathan", start: 690, end: 1110, comment: "" },
    { id: "seedp4-87", day: 1, employee: "Tom", start: 960, end: 1380, comment: "" },
    { id: "seedp4-88", day: 2, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-89", day: 2, employee: "Nathan", start: 690, end: 1110, comment: "" },
    { id: "seedp4-90", day: 2, employee: "Ewan", start: 810, end: 1380, comment: "" },
    { id: "seedp4-91", day: 3, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-92", day: 4, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-93", day: 4, employee: "Tom", start: 1020, end: 1380, comment: "" },
    { id: "seedp4-94", day: 5, employee: "Ewan", start: 570, end: 1230, comment: "" },
    { id: "seedp4-95", day: 6, employee: "Blanche", start: 570, end: 1230, comment: "" },
  ],
  "2026-12-21": [
    { id: "seedp4-96", day: 0, employee: "Aurélie", start: 1020, end: 1380, comment: "" },
    { id: "seedp4-97", day: 0, employee: "Nathan", start: 570, end: 990, comment: "" },
    { id: "seedp4-98", day: 0, employee: "Tom", start: 990, end: 1380, comment: "" },
    { id: "seedp4-99", day: 0, employee: "Ewan", start: 690, end: 1110, comment: "" },
    { id: "seedp4-100", day: 1, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-101", day: 1, employee: "Nathan", start: 810, end: 1380, comment: "" },
    { id: "seedp4-102", day: 1, employee: "Tom", start: 780, end: 1230, comment: "" },
    { id: "seedp4-103", day: 1, employee: "Ewan", start: 690, end: 1110, comment: "" },
    { id: "seedp4-104", day: 2, employee: "Aurélie", start: 1080, end: 1380, comment: "" },
    { id: "seedp4-105", day: 2, employee: "Tom", start: 690, end: 1260, comment: "" },
    { id: "seedp4-106", day: 2, employee: "Ewan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-107", day: 3, employee: "Aurélie", start: 570, end: 1110, comment: "" },
    { id: "seedp4-108", day: 5, employee: "Nathan", start: 570, end: 1230, comment: "" },
    { id: "seedp4-109", day: 6, employee: "Tom", start: 570, end: 1230, comment: "" },
  ],
  "2026-12-28": [
    { id: "seedp4-110", day: 0, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-111", day: 0, employee: "Nathan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-112", day: 0, employee: "Ewan", start: 570, end: 990, comment: "" },
    { id: "seedp4-113", day: 1, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-114", day: 1, employee: "Nathan", start: 990, end: 1380, comment: "" },
    { id: "seedp4-115", day: 1, employee: "Tom", start: 570, end: 990, comment: "" },
    { id: "seedp4-116", day: 1, employee: "Ewan", start: 570, end: 960, comment: "" },
    { id: "seedp4-117", day: 2, employee: "Aurélie", start: 990, end: 1380, comment: "" },
    { id: "seedp4-118", day: 2, employee: "Nathan", start: 690, end: 1110, comment: "" },
    { id: "seedp4-119", day: 2, employee: "Tom", start: 810, end: 1380, comment: "" },
    { id: "seedp4-120", day: 2, employee: "Ewan", start: 570, end: 900, comment: "" },
    { id: "seedp4-121", day: 3, employee: "Ewan", start: 570, end: 1110, comment: "" },
    { id: "seedp4-122", day: 5, employee: "Tom", start: 570, end: 1230, comment: "" },
    { id: "seedp4-123", day: 6, employee: "Blanche", start: 570, end: 1230, comment: "" },
  ],
};

const MARKERS_SEED = [
  { id: "mkseed-1", date: "2026-11-03", employee: "Aurélie", type: "cp" },
  { id: "mkseed-2", date: "2026-11-04", employee: "Aurélie", type: "cp" },
];


const EVENTS_SEED = {};


function empColor(name, employees) {
  const emp = employees.find((e) => e.name === name);
  if (!emp) return { bg: "bg-slate-400", text: "text-slate-600" };
  return EMP_PALETTE[emp.colorIdx % EMP_PALETTE.length];
}



const OUVERTURE_SEED = [
  { id: "o1", group: "Ouverture", text: "Allumer les lumières étagères, frigo, tireuse et salle", freq: "" },
  { id: "o2", group: "Ouverture", text: "Fermer la caisse de la veille (noter dans le fichier caisse)", freq: "" },
  { id: "o3", group: "Ouverture", text: "Multimédia : écran caméras + splitter + ampli + écrans TV (programme sports) + musique", freq: "" },
  { id: "o4", group: "Check", text: "État des sols : passage balai si besoin", freq: "" },
  { id: "o5", group: "Check", text: "État des tables : passage éponge si besoin", freq: "" },
  { id: "o6", group: "Check", text: "Mise en place de la salle (chaises, tabourets, tables)", freq: "" },
  { id: "o7", group: "Check", text: "Remplir les frigos + chips + snacking + saucissons", freq: "" },
  { id: "o8", group: "Sport", text: "Ouverture de la porte partie sport et allumer la musique", freq: "" },
  { id: "o9", group: "Sport", text: "Check état des terrains", freq: "" },
  { id: "o10", group: "Vestiaires & Sanitaire", text: "Check poubelles et affaires oubliées/douches", freq: "" },
  { id: "o11", group: "Vestiaires & Sanitaire", text: "Mettre un coup de sent-bon si besoin", freq: "" },
  { id: "o12", group: "Vestiaires & Sanitaire", text: "Check niveau PQ/essuie-mains/savons/désodorisant", freq: "" },
  { id: "o13", group: "Vestiaires & Sanitaire", text: "Check lave-linge", freq: "" },
];

const FERMETURE_SEED = [
  { id: "f1", group: "Bar & accueil", text: "Débarrasser et nettoyer les tables et le bar", freq: "" },
  { id: "f2", group: "Bar & accueil", text: "Nettoyer les sols (aspirateur et serpillère si besoin)", freq: "" },
  { id: "f3", group: "Bar & accueil", text: "Vaisselles/planches à laver/sécher/ranger", freq: "" },
  { id: "f4", group: "Bar & accueil", text: "Nettoyer la tireuse et la trancheuse", freq: "" },
  { id: "f5", group: "Bar & accueil", text: "Éteindre le lave-verres", freq: "" },
  { id: "f6", group: "Bar & accueil", text: "Check poubelles (remplacer si plus de la moitié)", freq: "" },
  { id: "f7", group: "Bar & accueil", text: "Faire le réassort", freq: "" },
  { id: "f8", group: "Bar & accueil", text: "Éteindre PC, TV, chauffage, lave-verre, fléchettes et lumières", freq: "" },
  { id: "f9", group: "Bar & accueil", text: "Fermer les fenêtres/portes", freq: "" },
  { id: "f10", group: "Salle Hybride", text: "Nettoyer les sols, check poubelle et produit DIY", freq: "" },
  { id: "f11", group: "Salle Hybride", text: "Éteindre chauffage", freq: "" },
  { id: "f12", group: "Cuisine", text: "Ranger et filmer toute la nourriture", freq: "" },
  { id: "f13", group: "Cuisine", text: "Nettoyer table et sol", freq: "" },
  { id: "f14", group: "Cuisine", text: "Prise T° frigo", freq: "", valueLabel: "Température (°C)" },
  { id: "f15", group: "Sport", text: "Check des terrains/check poubelles", freq: "" },
  { id: "f16", group: "Sport", text: "Éteindre ventilation/enceinte et la rentrer", freq: "" },
  { id: "f17", group: "Sport", text: "Fermer toutes les portes", freq: "" },
  { id: "f18", group: "Vestiaires & Sanitaire", text: "Check poubelles et affaires oubliées/douches", freq: "" },
  { id: "f19", group: "Vestiaires & Sanitaire", text: "Nettoyer les sols, WC et lavabo", freq: "" },
  { id: "f20", group: "Vestiaires & Sanitaire", text: "Check niveau PQ/essuie-mains/savons/désodorisant", freq: "" },
  { id: "f21", group: "Vestiaires & Sanitaire", text: "Lancer machine à laver avec chasubles sales", freq: "" },
  { id: "f22", group: "Sécurité & Clôture", text: "Fermer les 3 portes extérieures et mettre l'alarme", freq: "" },
  { id: "f23", group: "Sécurité & Clôture", text: "Vider cendriers et ranger les chaises/tables", freq: "" },
  { id: "f24", group: "Sécurité & Clôture", text: "Vérifier que le parking soit vide", freq: "" },
];

const ANNIV_SLOTS = ["13h30", "15h", "16h30"];
const ANNIV_DURATIONS = ["1h", "1h30"];
const FFT_LEVELS = ["P25", "P50", "P100", "P250", "P500"];
const LOISIRS_TYPES = ["Tournoi padel loisirs", "Tournoi squash", "Tournoi badminton", "Tournoi pickleball", "Tournoi 4 raquettes", "Autres"];

const EVENT_GROUPS_SEED = [
  { id: "eg1", name: "Event" },
  { id: "eg3", name: "Actualité" },
];

const EVENT_ENTRIES_SEED = [
  { id: "evtseed-1", date: "2026-08-31", kind: "simple", category: "Actualité", comment: "Installation FOOT" },
  { id: "evtseed-2", date: "2026-09-01", kind: "simple", category: "Actualité", comment: "Installation FOOT" },
  { id: "evtseed-3", date: "2026-09-02", kind: "simple", category: "Actualité", comment: "Installation FOOT" },
  { id: "evtseed-4", date: "2026-09-03", kind: "tournoi", subtype: "fft", level: "P100", comment: "" },
  { id: "evtseed-5", date: "2026-09-03", kind: "simple", category: "Actualité", comment: "Installation FOOT" },
  { id: "evtseed-6", date: "2026-09-04", kind: "reservation", forWhom: "10H - 11H30 : Padel Business Club", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-7", date: "2026-09-04", kind: "simple", category: "Actualité", comment: "Installation FOOT" },
  { id: "evtseed-8", date: "2026-09-06", kind: "simple", category: "Event", comment: "Tournoi Padel Découverte Femmes + Cours Yoga : 10h - 12h" },
  { id: "evtseed-9", date: "2026-09-09", kind: "tournoi", subtype: "fft", level: "P50", comment: "" },
  { id: "evtseed-10", date: "2026-09-10", kind: "reservation", forWhom: "EPIDE 10h - 11h30", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-11", date: "2026-09-11", kind: "simple", category: "Event", comment: "Leizup" },
  { id: "evtseed-12", date: "2026-09-11", kind: "tournoi", subtype: "loisirs", loisirsType: "Autres", comment: "Montante/Descendante" },
  { id: "evtseed-13", date: "2026-09-12", kind: "simple", category: "Event", comment: "STAND Décathlon" },
  { id: "evtseed-14", date: "2026-09-13", kind: "simple", category: "Event", comment: "Ambulances Jussieu" },
  { id: "evtseed-15", date: "2026-09-15", kind: "reservation", forWhom: "Bar: deci'apero", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-16", date: "2026-09-18", kind: "simple", category: "Event", comment: "COX" },
  { id: "evtseed-17", date: "2026-09-18", kind: "tournoi", subtype: "fft", level: "P50", comment: "" },
  { id: "evtseed-18", date: "2026-09-19", kind: "tournoi", subtype: "fft", level: "P100", comment: "" },
  { id: "evtseed-19", date: "2026-09-24", kind: "reservation", forWhom: "Tournoi Infinity - 9h30 - 12h30", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-20", date: "2026-09-24", kind: "tournoi", subtype: "loisirs", loisirsType: "Autres", comment: "Montante/Descendante" },
  { id: "evtseed-21", date: "2026-09-25", kind: "simple", category: "Event", comment: "Comedy Club" },
  { id: "evtseed-22", date: "2026-09-26", kind: "tournoi", subtype: "loisirs", loisirsType: "Tournoi padel loisirs", comment: "" },
  { id: "evtseed-23", date: "2026-09-29", kind: "tournoi", subtype: "fft", level: "P25", comment: "" },
  { id: "evtseed-24", date: "2026-10-02", kind: "reservation", forWhom: "10H - 11H30 : Padel Business Club", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-25", date: "2026-10-02", kind: "tournoi", subtype: "loisirs", loisirsType: "Tournoi squash", comment: "" },
  { id: "evtseed-26", date: "2026-10-03", kind: "tournoi", subtype: "fft", level: "P100", comment: "" },
  { id: "evtseed-27", date: "2026-10-07", kind: "tournoi", subtype: "fft", level: "P50", comment: "" },
  { id: "evtseed-28", date: "2026-10-08", kind: "reservation", forWhom: "Bar: deci'dej", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-29", date: "2026-10-08", kind: "tournoi", subtype: "loisirs", loisirsType: "Autres", comment: "Montante/Descendante" },
  { id: "evtseed-30", date: "2026-10-09", kind: "reservation", forWhom: "Bar et Sport: OPELLA - 100 personnes", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-31", date: "2026-10-10", kind: "tournoi", subtype: "fft", level: "P250", comment: "" },
  { id: "evtseed-32", date: "2026-10-13", kind: "reservation", forWhom: "EPIDE 10h - 11h30", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-33", date: "2026-10-13", kind: "tournoi", subtype: "fft", level: "P50", comment: "" },
  { id: "evtseed-34", date: "2026-10-15", kind: "reservation", forWhom: "Soirée Projection Mathieu Durieux - 19h", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-35", date: "2026-10-16", kind: "simple", category: "Event", comment: "COX" },
  { id: "evtseed-36", date: "2026-10-17", kind: "tournoi", subtype: "fft", level: "P100", comment: "" },
  { id: "evtseed-37", date: "2026-10-22", kind: "tournoi", subtype: "fft", level: "P100", comment: "" },
  { id: "evtseed-38", date: "2026-10-23", kind: "tournoi", subtype: "loisirs", loisirsType: "Autres", comment: "Montante/Descendante" },
  { id: "evtseed-39", date: "2026-10-29", kind: "tournoi", subtype: "fft", level: "P100", comment: "" },
  { id: "evtseed-40", date: "2026-10-30", kind: "tournoi", subtype: "loisirs", loisirsType: "Tournoi pickleball", comment: "" },
  { id: "evtseed-41", date: "2026-10-31", kind: "simple", category: "Event", comment: "Soirée Delvallez ? 70PAX" },
  { id: "evtseed-42", date: "2026-10-31", kind: "tournoi", subtype: "fft", level: "P50", comment: "" },
  { id: "evtseed-43", date: "2026-11-03", kind: "tournoi", subtype: "fft", level: "P100", comment: "" },
  { id: "evtseed-44", date: "2026-11-05", kind: "reservation", forWhom: "Bar: deci'apero", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-45", date: "2026-11-05", kind: "tournoi", subtype: "loisirs", loisirsType: "Autres", comment: "Montante/Descendante" },
  { id: "evtseed-46", date: "2026-11-06", kind: "reservation", forWhom: "10H - 11H30 : Padel Business Club", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-47", date: "2026-11-12", kind: "tournoi", subtype: "fft", level: "P100", comment: "" },
  { id: "evtseed-48", date: "2026-11-13", kind: "tournoi", subtype: "loisirs", loisirsType: "Tournoi badminton", comment: "" },
  { id: "evtseed-49", date: "2026-11-14", kind: "tournoi", subtype: "fft", level: "P250", comment: "" },
  { id: "evtseed-50", date: "2026-11-17", kind: "reservation", forWhom: "EPIDE 10h - 11h30", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-51", date: "2026-11-19", kind: "tournoi", subtype: "loisirs", loisirsType: "Autres", comment: "Montante/Descendante" },
  { id: "evtseed-52", date: "2026-11-20", kind: "simple", category: "Event", comment: "Comedy Club" },
  { id: "evtseed-53", date: "2026-11-21", kind: "tournoi", subtype: "fft", level: "P25", comment: "" },
  { id: "evtseed-54", date: "2026-11-26", kind: "reservation", forWhom: "Bar: deci'dej", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-55", date: "2026-11-26", kind: "tournoi", subtype: "fft", level: "P100", comment: "" },
  { id: "evtseed-56", date: "2026-11-27", kind: "simple", category: "Event", comment: "COX" },
  { id: "evtseed-57", date: "2026-11-27", kind: "tournoi", subtype: "loisirs", loisirsType: "Tournoi squash", comment: "" },
  { id: "evtseed-58", date: "2026-11-28", kind: "tournoi", subtype: "fft", level: "P500", comment: "" },
  { id: "evtseed-59", date: "2026-12-02", kind: "tournoi", subtype: "fft", level: "P50", comment: "" },
  { id: "evtseed-60", date: "2026-12-04", kind: "tournoi", subtype: "loisirs", loisirsType: "Tournoi 4 raquettes", comment: "" },
  { id: "evtseed-61", date: "2026-12-05", kind: "tournoi", subtype: "fft", level: "P100", comment: "" },
  { id: "evtseed-62", date: "2026-12-10", kind: "reservation", forWhom: "EPIDE 10h - 11h30", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-63", date: "2026-12-11", kind: "tournoi", subtype: "loisirs", loisirsType: "Autres", comment: "Montante/Descendante" },
  { id: "evtseed-64", date: "2026-12-12", kind: "tournoi", subtype: "fft", level: "P100", comment: "" },
  { id: "evtseed-65", date: "2026-12-18", kind: "simple", category: "Event", comment: "COX" },
  { id: "evtseed-66", date: "2026-12-18", kind: "tournoi", subtype: "loisirs", loisirsType: "Autres", comment: "montente descendante 4 ans du stadium" },
  { id: "evtseed-67", date: "2026-12-22", kind: "reservation", forWhom: "Bar: deci'apero", isBar: false, isSport: false, time: "", description: "", confirmed: false },
  { id: "evtseed-68", date: "2026-12-22", kind: "tournoi", subtype: "fft", level: "P100", comment: "" },
  { id: "evtseed-69", date: "2026-12-29", kind: "tournoi", subtype: "fft", level: "P100", comment: "" },
];

function hasAssignee(item, name) {
  const list = item.assignees || (item.assignee ? [item.assignee] : []);
  return list.includes(name) || list.includes("Tout le monde");
}

function summarizeEntry(e) {
  if (e.kind === "anniversaire") {
    const name = [e.firstName, e.lastName].filter(Boolean).join(" ");
    return "Anniversaire: " + (name || "?") + (e.slot ? " à " + e.slot : "");
  }
  if (e.kind === "reservation") {
    const types = [e.isBar && "bar", e.isSport && "sport"].filter(Boolean).join(" et ");
    return "Réservation: " + (e.forWhom || "?") + (types ? " " + types : "") + (e.time ? " à " + e.time : "");
  }
  if (e.kind === "tournoi") {
    if (e.subtype === "fft") {
      return "Tournoi FFT " + (e.level || "?") + (e.comment ? ": " + e.comment : "");
    }
    const label = e.loisirsType === "Autres" ? "Autres" + (e.comment ? " (" + e.comment + ")" : "") : e.loisirsType || "?";
    return "Tournoi loisirs: " + label;
  }
  return (e.category || "Évènement") + ": " + (e.comment || "");
}

const PRIORITIES = [
  { key: "haute", label: "Haute", bg: "bg-rose-100", text: "text-rose-700" },
  { key: "moyenne", label: "Moyenne", bg: "bg-amber-100", text: "text-amber-700" },
  { key: "faible", label: "Faible", bg: "bg-slate-100", text: "text-slate-500" },
];
const PRIORITY_RANK = { haute: 0, moyenne: 1, faible: 2 };

const ACTIONS_SEED = [
  { id: "a1", text: "Passer l'autolaveuse", freq: "Tous les dimanches" },
  { id: "a2", text: "Aspirateur squash & badminton", freq: "Mercredi, une semaine sur 2" },
  { id: "a3", text: "Ménage complet de l'entrepôt", freq: "1er mercredi tous les 2 mois" },
  { id: "a4", text: "Nettoyage des sols vestiaires au karcher", freq: "1er mardi du mois" },
  { id: "a5", text: "Nettoyage des parois vestiaire à l'anticalcaire", freq: "1er mardi du mois" },
  { id: "a6", text: "Check inventaire produit ménager", freq: "1er lundi du mois" },
  { id: "a7", text: "Rangement buanderie", freq: "1er lundi, tous les 2 mois" },
  { id: "a8", text: "Check niveau sel dans l'adoucisseur", freq: "Lundi, une semaine sur 2" },
  { id: "a9", text: "Check date de péremption étagère", freq: "1er lundi du mois" },
  { id: "a10", text: "Ménage complet de la cuisine", freq: "1er lundi du mois" },
].map((a) => ({ ...a, date: "", assignee: "", comment: "", done: false, doneBy: null, doneTs: null }));

const STOCK_SEED = [
  { id: "s1", group: "Bar", name: "Bière blonde" },
  { id: "s2", group: "Bar", name: "Vedett IPA" },
  { id: "s3", group: "Bar", name: "Filou" },
  { id: "s4", group: "Bar", name: "Kastel rouge" },
  { id: "s5", group: "Bar", name: "Desperados" },
  { id: "s6", group: "Bar", name: "Vin rouge" },
  { id: "s7", group: "Bar", name: "Vin blanc" },
  { id: "s8", group: "Bar", name: "Vin rosé" },
  { id: "s9", group: "Bar", name: "Jus orange" },
  { id: "s10", group: "Bar", name: "Jus ACE" },
  { id: "s11", group: "Bar", name: "Jus ananas" },
  { id: "s12", group: "Bar", name: "Limonade" },
  { id: "s13", group: "Bar", name: "Coca" },
  { id: "s14", group: "Bar", name: "Coca Zéro" },
  { id: "s15", group: "Bar", name: "Orangina" },
  { id: "s16", group: "Bar", name: "IceTea" },
  { id: "s17", group: "Bar", name: "Powerade" },
  { id: "s18", group: "Bar", name: "Red Bull" },
  { id: "s19", group: "Bar", name: "Sprite" },
  { id: "s20", group: "Bar", name: "Oasis" },
  { id: "s21", group: "Bar", name: "Bouteille d'eau" },
  { id: "s22", group: "Bar", name: "Limonade bouteille" },
  { id: "s23", group: "Bar", name: "Picon" },
  { id: "s24", group: "Cuisine", name: "Comté" },
  { id: "s25", group: "Cuisine", name: "Cabecou" },
  { id: "s26", group: "Cuisine", name: "Coppa" },
  { id: "s27", group: "Cuisine", name: "Spianata" },
  { id: "s28", group: "Cuisine", name: "Porceta" },
  { id: "s29", group: "Cuisine", name: "Saucisson" },
  { id: "s30", group: "Cuisine", name: "Jambon blanc" },
  { id: "s31", group: "Cuisine", name: "Pizza" },
  { id: "s32", group: "Cuisine", name: "Cossette mozza" },
  { id: "s33", group: "Cuisine", name: "Roquefort" },
  { id: "s34", group: "Cuisine", name: "Chèvre" },
  { id: "s35", group: "Cuisine", name: "Beurre micropain" },
  { id: "s36", group: "Cuisine", name: "Beurre" },
  { id: "s37", group: "Cuisine", name: "Boîte pizza" },
  { id: "s38", group: "Cuisine", name: "Origan" },
  { id: "s39", group: "Cuisine", name: "Cornichon" },
  { id: "s40", group: "Cuisine", name: "Olive" },
  { id: "s41", group: "Cuisine", name: "Champignon" },
  { id: "s42", group: "Cuisine", name: "Pain panini" },
  { id: "s43", group: "Cuisine", name: "Chips" },
  { id: "s44", group: "Cuisine", name: "Bonbon" },
];

function pct(minutes) {
  return Math.max(0, Math.min(100, (minutes / RANGE_TOTAL) * 100));
}
function fmtTime(m) {
  const h = Math.floor(m / 60);
  const mm = m % 60;
  return mm === 0 ? `${h}h` : `${h}h${String(mm).padStart(2, "0")}`;
}
function mondayOf(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = (day === 0 ? -6 : 1) - day;
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}
function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}
function isoDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}
function fmtDayLabel(d) {
  const s = d.toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long" });
  return s.charAt(0).toUpperCase() + s.slice(1);
}
function fmtShort(d) {
  return d.toLocaleDateString("fr-FR", { day: "numeric", month: "short" });
}
function sameDay(a, b) {
  return isoDate(a) === isoDate(b);
}
async function storageGet(key, fallback) {
  try {
    const { data, error } = await supabase.from("app_storage").select("value").eq("key", key).maybeSingle();
    if (error || !data) return fallback;
    return data.value;
  } catch (e) {
    return fallback;
  }
}
async function storageSet(key, value) {
  try {
    await supabase.from("app_storage").upsert({ key, value, updated_at: new Date().toISOString() });
  } catch (e) {
    // best effort
  }
}
async function storageDeleteByPrefix(prefix) {
  try {
    await supabase.from("app_storage").delete().like("key", prefix + "%");
  } catch (e) {
    // best effort
  }
}
async function storageListByPrefix(prefix) {
  try {
    const { data, error } = await supabase.from("app_storage").select("key, value").like("key", prefix + "%");
    if (error || !data) return [];
    return data;
  } catch (e) {
    return [];
  }
}

export default function App() {
  const [profile, setProfile] = useState(null);
  const [profileOpen, setProfileOpen] = useState(false);
  const [profileType, setProfileType] = useState(null);
  const [managerSession, setManagerSession] = useState(null);
  const [loginEmail, setLoginEmail] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [loginLoading, setLoginLoading] = useState(false);
  const [now, setNow] = useState(new Date());
  const [actions, setActions] = useState([]);
  const [managerActions, setManagerActions] = useState([]);
  const [projects, setProjects] = useState([]);
  const [employees, setEmployees] = useState([]);
  const [markers, setMarkers] = useState([]);
  const [stockProducts, setStockProducts] = useState([]);
  const [stockChecked, setStockChecked] = useState([]);
  const [eventEntries, setEventEntries] = useState([]);
  const [eventGroups, setEventGroups] = useState([]);
  const [activeTab, setActiveTab] = useState("planning");

  useEffect(() => {
    const t = setInterval(() => setNow(new Date()), 1000 * 15);
    return () => clearInterval(t);
  }, []);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setManagerSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, session) => {
      setManagerSession(session);
    });
    return () => sub.subscription.unsubscribe();
  }, []);

  async function handleManagerLogin(e) {
    e.preventDefault();
    setLoginError("");
    setLoginLoading(true);
    const { error } = await supabase.auth.signInWithPassword({ email: loginEmail, password: loginPassword });
    setLoginLoading(false);
    if (error) {
      setLoginError("Email ou mot de passe incorrect.");
      return;
    }
    setLoginPassword("");
  }

  async function handleManagerLogout() {
    await supabase.auth.signOut();
    setProfile(null);
    setProfileType(null);
    setProfileOpen(false);
  }

  useEffect(() => {
    (async () => {
      let a = await storageGet("actions-list", null);
      if (a === null) {
        a = ACTIONS_SEED;
        await storageSet("actions-list", a);
      }
      setActions(a);
    })();
    storageGet("manager-actions-list", []).then(setManagerActions);
    storageGet("projects-list", []).then(setProjects);
    (async () => {
      let e = await storageGet("employees:v2", null);
      if (e === null) {
        e = EMPLOYEES_SEED;
        await storageSet("employees:v2", e);
      }
      setEmployees(e);
    })();
    (async () => {
      let m = await storageGet("day-markers:v2", null);
      if (m === null) {
        m = MARKERS_SEED;
        await storageSet("day-markers:v2", m);
      }
      setMarkers(m);
    })();
    (async () => {
      let sp = await storageGet("stock-products:v2", null);
      if (sp === null) {
        sp = STOCK_SEED;
        await storageSet("stock-products:v2", sp);
      }
      setStockProducts(sp);
    })();
    storageGet("stock-checked", []).then(setStockChecked);
    (async () => {
      let ee = await storageGet("event-entries:v2", null);
      if (ee === null) {
        ee = EVENT_ENTRIES_SEED;
        await storageSet("event-entries:v2", ee);
      }
      setEventEntries(ee);
    })();
    (async () => {
      let eg = await storageGet("event-groups:v2", null);
      if (eg === null) {
        eg = EVENT_GROUPS_SEED;
        await storageSet("event-groups:v2", eg);
      }
      setEventGroups(eg);
    })();
  }, []);

  const saveActions = useCallback(async (next) => {
    setActions(next);
    await storageSet("actions-list", next);
  }, []);

  const saveManagerActions = useCallback(async (next) => {
    setManagerActions(next);
    await storageSet("manager-actions-list", next);
  }, []);

  const saveProjects = useCallback(async (next) => {
    setProjects(next);
    await storageSet("projects-list", next);
  }, []);

  const saveEmployees = useCallback(async (next) => {
    setEmployees(next);
    await storageSet("employees:v2", next);
  }, []);

  const saveMarkers = useCallback(async (next) => {
    setMarkers(next);
    await storageSet("day-markers:v2", next);
  }, []);

  const saveStockProducts = useCallback(async (next) => {
    setStockProducts(next);
    await storageSet("stock-products:v2", next);
  }, []);

  const saveStockChecked = useCallback(async (next) => {
    setStockChecked(next);
    await storageSet("stock-checked", next);
  }, []);

  const saveEventEntries = useCallback(async (next) => {
    setEventEntries(next);
    await storageSet("event-entries:v2", next);
  }, []);

  const saveEventGroups = useCallback(async (next) => {
    setEventGroups(next);
    await storageSet("event-groups:v2", next);
  }, []);

  function selectProfile(name, role) {
    setProfile({ name, role });
    setProfileOpen(false);
    setProfileType(null);
    if (role !== "manager" && activeTab === "reglages") setActiveTab("planning");
  }

  const isManager = profile && profile.role === "manager";
  const dateLabel = now.toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long" });
  const oswald = { fontFamily: "'Oswald', sans-serif" };

  return (
    <div className="min-h-screen bg-slate-100">
      <style>{"@import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;700&display=swap');"}</style>

      <header className="px-4 py-3" style={{ background: "linear-gradient(135deg, #064e3b, #022c22)" }}>
        <div className="max-w-5xl mx-auto flex items-center justify-between">
          <div className="relative">
            <button onClick={() => { setProfileOpen((o) => !o); setProfileType(null); }} className={"px-3 py-2 rounded-lg text-sm font-bold flex items-center gap-1.5 " + (profile ? "bg-emerald-800 text-white border border-emerald-700" : "bg-amber-400 text-emerald-900")}>
              {profile ? (
                <span className="text-left leading-tight">
                  <span className="block">{profile.name}</span>
                  <span className="block text-xs font-medium capitalize opacity-80">{profile.role === "manager" ? "Manager" : "Employé"}</span>
                </span>
              ) : (
                <span>Sélectionner un profil</span>
              )}
              <ChevronDown size={16} />
            </button>
            {profileOpen && (
              <>
                <div className="fixed inset-0 z-40" onClick={() => { setProfileOpen(false); setProfileType(null); }} />
                <div className="absolute left-0 mt-2 w-64 bg-white rounded-xl shadow-lg border border-slate-200 z-50 overflow-hidden">
                  {!profileType ? (
                    <div className="p-2">
                      <button onClick={() => setProfileType("manager")} className="w-full text-left px-3 py-3 rounded-lg hover:bg-emerald-50 text-slate-800 font-semibold">
                        Manager
                      </button>
                      <button onClick={() => setProfileType("employe")} className="w-full text-left px-3 py-3 rounded-lg hover:bg-emerald-50 text-slate-800 font-semibold">
                        Employé
                      </button>
                    </div>
                  ) : profileType === "manager" && !managerSession ? (
                    <div>
                      <button onClick={() => setProfileType(null)} className="w-full text-left px-4 py-2 text-xs font-bold text-slate-400 uppercase tracking-wide hover:bg-slate-50">
                        ← Manager
                      </button>
                      <form onSubmit={handleManagerLogin} className="p-3 pt-1">
                        <label className="text-xs font-semibold text-slate-500 mb-1 block">Email</label>
                        <input type="email" required value={loginEmail} onChange={(e) => setLoginEmail(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-2" />
                        <label className="text-xs font-semibold text-slate-500 mb-1 block">Mot de passe</label>
                        <input type="password" required value={loginPassword} onChange={(e) => setLoginPassword(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-2" />
                        {loginError && <div className="text-rose-600 text-xs mb-2">{loginError}</div>}
                        <button type="submit" disabled={loginLoading} className="w-full py-2 rounded-lg bg-emerald-600 text-white font-bold text-sm">
                          {loginLoading ? "Connexion…" : "Se connecter"}
                        </button>
                      </form>
                    </div>
                  ) : (
                    <div>
                      <button onClick={() => setProfileType(null)} className="w-full text-left px-4 py-2 text-xs font-bold text-slate-400 uppercase tracking-wide hover:bg-slate-50">
                        ← {profileType === "manager" ? "Manager" : "Employé"}
                      </button>
                      {(profileType === "manager" ? MANAGERS : employees.map((e) => e.name)).map((name) => (
                        <button key={name} onClick={() => selectProfile(name, profileType)} className="w-full text-left px-4 py-2 hover:bg-emerald-50 text-slate-800 font-medium border-t border-slate-100">
                          {name}
                        </button>
                      ))}
                      {profileType === "manager" && (
                        <button onClick={handleManagerLogout} className="w-full text-left px-4 py-2 text-xs text-slate-400 border-t border-slate-100 hover:bg-slate-50">
                          Déconnexion manager
                        </button>
                      )}
                    </div>
                  )}
                </div>
              </>
            )}
          </div>

          <div className="flex items-center gap-3">
            <div className="text-right">
              <div className="text-white font-bold tracking-wide leading-none" style={oswald}>LE STADIUM</div>
              <div className="text-emerald-300 text-xs uppercase tracking-widest mt-1">Team</div>
            </div>
            <img src={LOGO_B64} alt="Le Stadium" style={{ height: "40px", width: "auto" }} />
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-3 sm:px-4 py-5">
        {!profile && (
          <div className="text-center py-16">
            <div className="w-16 h-16 rounded-2xl bg-emerald-100 text-emerald-700 flex items-center justify-center mx-auto mb-4">
              <Users size={28} />
            </div>
            <div className="text-lg font-bold text-emerald-900" style={oswald}>Bienvenue</div>
            <p className="text-slate-500 mt-1">Sélectionnez votre profil en haut à gauche pour commencer.</p>
          </div>
        )}

        {profile && (
          <AppBody profile={profile} isManager={isManager} activeTab={activeTab} setActiveTab={setActiveTab} oswald={oswald} actions={actions} saveActions={saveActions} managerActions={managerActions} saveManagerActions={saveManagerActions} projects={projects} saveProjects={saveProjects} employees={employees} saveEmployees={saveEmployees} markers={markers} saveMarkers={saveMarkers} stockProducts={stockProducts} saveStockProducts={saveStockProducts} stockChecked={stockChecked} saveStockChecked={saveStockChecked} eventEntries={eventEntries} saveEventEntries={saveEventEntries} eventGroups={eventGroups} saveEventGroups={saveEventGroups} />
        )}
      </main>
    </div>
  );
}

function AppBody({ profile, isManager, activeTab, setActiveTab, oswald, actions, saveActions, managerActions, saveManagerActions, projects, saveProjects, employees, saveEmployees, markers, saveMarkers, stockProducts, saveStockProducts, stockChecked, saveStockChecked, eventEntries, saveEventEntries, eventGroups, saveEventGroups }) {
  const todayISO = isoDate(new Date());
  const pendingOwnActions = actions.filter((a) => !a.done && hasAssignee(a, profile.name)).length;
  const pendingManagerActions = isManager ? managerActions.filter((a) => !a.done && hasAssignee(a, profile.name)).length : 0;
  const overdueProjects = projects.filter((p) => p.dueDate && p.status !== "termine" && p.dueDate < todayISO && hasAssignee(p, profile.name)).length;
  const pendingActions = pendingOwnActions + pendingManagerActions + overdueProjects;
  const tabs = [
    { key: "planning", label: "Planning" },
    { key: "evenement", label: "Événement" },
    { key: "ouverture", label: "Ouverture" },
    { key: "fermeture", label: "Fermeture" },
    { key: "todo", label: "To do", badge: pendingActions },
    { key: "stock", label: "Stock", badge: stockChecked.length },
  ];
  if (isManager) {
    tabs.push({ key: "cp", label: "CP" });
    tabs.push({ key: "historique", label: "Historique" });
    tabs.push({ key: "reglages", label: "Réglages" });
  }

  return (
    <div>
      <div className="flex gap-1 mb-4 border-b border-slate-200 overflow-x-auto whitespace-nowrap">
        {tabs.map((t) => (
          <button
            key={t.key}
            onClick={() => setActiveTab(t.key)}
            className={"relative px-3 py-2 font-semibold text-sm border-b-2 -mb-px shrink-0 " + (activeTab === t.key ? "border-emerald-600 text-emerald-800" : "border-transparent text-slate-400")}
            style={oswald}
          >
            {t.label}
            {t.badge > 0 && (
              <span className="absolute -top-0.5 -right-2 bg-rose-500 text-white rounded-full flex items-center justify-center font-bold" style={{ width: "16px", height: "16px", fontSize: "9px" }}>
                {t.badge}
              </span>
            )}
          </button>
        ))}
      </div>

      {activeTab === "planning" && <PlanningTab profile={profile} isManager={isManager} oswald={oswald} employees={employees} markers={markers} saveMarkers={saveMarkers} eventEntries={eventEntries} />}
      {activeTab === "evenement" && <EvenementTab profile={profile} eventEntries={eventEntries} saveEventEntries={saveEventEntries} eventGroups={eventGroups} oswald={oswald} />}
      {activeTab === "ouverture" && <ChecklistTab type="ouverture" title="Process ouverture" seed={OUVERTURE_SEED} isManager={isManager} profile={profile} oswald={oswald} />}
      {activeTab === "fermeture" && <ChecklistTab type="fermeture" title="Process fermeture" seed={FERMETURE_SEED} isManager={isManager} profile={profile} oswald={oswald} />}
      {activeTab === "todo" && <TodoTab profile={profile} isManager={isManager} actions={actions} saveActions={saveActions} managerActions={managerActions} saveManagerActions={saveManagerActions} projects={projects} saveProjects={saveProjects} employees={employees} oswald={oswald} />}
      {activeTab === "stock" && <StockTab products={stockProducts} saveProducts={saveStockProducts} checked={stockChecked} saveChecked={saveStockChecked} oswald={oswald} />}
      {activeTab === "cp" && isManager && <CPTab markers={markers} employees={employees} oswald={oswald} />}
      {activeTab === "historique" && isManager && <HistoriqueTab oswald={oswald} />}
      {activeTab === "reglages" && isManager && <SettingsTab employees={employees} saveEmployees={saveEmployees} eventGroups={eventGroups} saveEventGroups={saveEventGroups} saveActions={saveActions} saveEventEntries={saveEventEntries} oswald={oswald} />}
    </div>
  );
}

/* ---------------- PLANNING ---------------- */

function PlanningTab({ profile, isManager, oswald, employees, markers, saveMarkers, eventEntries }) {
  const [weekStart, setWeekStart] = useState(mondayOf(new Date()));
  const [slots, setSlots] = useState([]);
  const [events, setEvents] = useState({});
  const [loading, setLoading] = useState(true);
  const [modal, setModal] = useState(null);
  const [viewSlot, setViewSlot] = useState(null);
  const [view, setView] = useState("detail");
  const [weekPickerOpen, setWeekPickerOpen] = useState(false);

  const weekKey = isoDate(weekStart);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    (async () => {
      let s = await storageGet("planning:" + weekKey, null);
      if (s === null) {
        s = PLANNING_SEED[weekKey] || [];
        await storageSet("planning:" + weekKey, s);
      }
      let e = await storageGet("events:" + weekKey, null);
      if (e === null) {
        e = EVENTS_SEED[weekKey] || {};
        await storageSet("events:" + weekKey, e);
      }
      if (!cancelled) {
        setSlots(s);
        setEvents(e);
        setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [weekKey]);

  const saveSlots = useCallback(async (next) => {
    setSlots(next);
    await storageSet("planning:" + weekKey, next);
  }, [weekKey]);

  const saveEvents = useCallback(async (next) => {
    setEvents(next);
    await storageSet("events:" + weekKey, next);
  }, [weekKey]);

  function changeWeek(delta) {
    setWeekStart((w) => addDays(w, delta * 7));
  }
  function goToday() {
    setWeekStart(mondayOf(new Date()));
  }

  function openAddModal(dayIndex, employee, x, width) {
    const ratio = Math.min(1, Math.max(0, x / width));
    let startMin = RANGE_START + ratio * RANGE_TOTAL;
    startMin = Math.round(startMin / 30) * 30;
    startMin = Math.min(startMin, RANGE_END - 60);
    const endMin = Math.min(startMin + 120, RANGE_END);
    setModal({ mode: "add", id: null, day: dayIndex, employee, start: startMin, end: endMin, comment: "" });
  }
  function openEditModal(slot) {
    if (isManager) setModal({ mode: "edit", ...slot });
    else setViewSlot(slot);
  }
  async function handleSaveSlot(data) {
    if (data.end <= data.start) return;
    let next;
    if (modal.mode === "add") {
      const id = Date.now() + "-" + Math.random().toString(36).slice(2, 8);
      next = [...slots, { id, day: modal.day, employee: modal.employee, start: data.start, end: data.end, comment: data.comment }];
    } else {
      next = slots.map((s) => (s.id === modal.id ? { ...s, start: data.start, end: data.end, comment: data.comment } : s));
    }
    await saveSlots(next);
    setModal(null);
  }
  async function handleDeleteSlot() {
    await saveSlots(slots.filter((s) => s.id !== modal.id));
    setModal(null);
  }
  async function handleEventChange(dayIndex, text) {
    await saveEvents({ ...events, [dayIndex]: text });
  }

  function markerFor(dateISO, employee) {
    const m = markers.find((mk) => mk.date === dateISO && mk.employee === employee);
    return m ? m.type : null;
  }
  async function toggleMarker(dateISO, employee, type) {
    const idx = markers.findIndex((mk) => mk.date === dateISO && mk.employee === employee);
    let next;
    if (type === null) {
      next = markers.filter((_, i) => i !== idx);
    } else if (idx >= 0) {
      next = markers.map((mk, i) => (i === idx ? { ...mk, type } : mk));
    } else {
      next = [...markers, { id: "mk-" + Date.now() + "-" + Math.random().toString(36).slice(2, 6), date: dateISO, employee, type }];
    }
    await saveMarkers(next);
  }

  const weekEnd = addDays(weekStart, 6);

  return (
    <div>
      <div className="flex items-center justify-between mb-3">
        <button onClick={() => changeWeek(-1)} className="p-2 rounded-lg bg-white shadow-sm border border-slate-200 text-slate-600">
          <ChevronLeft size={18} />
        </button>
        <div className="text-center relative">
          <button onClick={() => setWeekPickerOpen((o) => !o)} className="font-bold text-emerald-900" style={{ ...oswald, textDecoration: "underline", textDecorationStyle: "dotted" }}>
            Semaine du {fmtShort(weekStart)} au {fmtShort(weekEnd)}
          </button>
          <div>
            <button onClick={goToday} className="text-xs text-amber-600 font-semibold">Revenir à aujourd'hui</button>
          </div>
          {weekPickerOpen && (
            <>
              <div className="fixed inset-0 z-40" onClick={() => setWeekPickerOpen(false)} />
              <div className="absolute mt-2 z-50 bg-white rounded-lg shadow-lg border border-slate-200 p-3" style={{ left: "50%", transform: "translateX(-50%)" }}>
                <input
                  type="date"
                  autoFocus
                  onChange={(e) => {
                    if (e.target.value) {
                      setWeekStart(mondayOf(new Date(e.target.value + "T00:00:00")));
                      setWeekPickerOpen(false);
                    }
                  }}
                  className="border border-slate-200 rounded-lg px-2 py-1.5 text-sm"
                />
              </div>
            </>
          )}
        </div>
        <button onClick={() => changeWeek(1)} className="p-2 rounded-lg bg-white shadow-sm border border-slate-200 text-slate-600">
          <ChevronRight size={18} />
        </button>
      </div>

      <div className="flex items-center justify-between mb-3">
        <div className="flex bg-white rounded-lg border border-slate-200 p-0.5">
          <button
            onClick={() => setView("detail")}
            className={"px-3 py-1.5 rounded-md text-xs font-bold flex items-center gap-1 " + (view === "detail" ? "bg-emerald-600 text-white" : "text-slate-500")}
          >
            <LayoutGrid size={14} /> Détail
          </button>
          <button
            onClick={() => setView("recap")}
            className={"px-3 py-1.5 rounded-md text-xs font-bold flex items-center gap-1 " + (view === "recap" ? "bg-emerald-600 text-white" : "text-slate-500")}
          >
            <Table2 size={14} /> Récap
          </button>
        </div>
        {isManager && view === "detail" && (
          <p className="text-xs text-slate-400 hidden sm:block">Touchez une ligne pour ajouter, un créneau pour le modifier.</p>
        )}
      </div>

      {loading ? (
        <div className="text-center text-slate-400 py-10">Chargement du planning…</div>
      ) : view === "recap" ? (
        <RecapView weekStart={weekStart} slots={slots} oswald={oswald} />
      ) : (
        Array.from({ length: 7 }, (_, i) => addDays(weekStart, i)).map((date, dayIndex) => (
          <DayCard
            key={dayIndex}
            date={date}
            dayIndex={dayIndex}
            slots={slots.filter((s) => s.day === dayIndex)}
            isManager={isManager}
            onAddSlot={openAddModal}
            onEditSlot={openEditModal}
            currentEmployeeName={profile.role === "employe" ? profile.name : null}
            oswald={oswald}
            eventText={events[dayIndex] || ""}
            onEventChange={(text) => handleEventChange(dayIndex, text)}
            employees={employees}
            markerFor={(employee) => markerFor(isoDate(date), employee)}
            onToggleMarker={(employee, type) => toggleMarker(isoDate(date), employee, type)}
            dayEvents={eventEntries.filter((e) => e.date === isoDate(date))}
          />
        ))
      )}

      {!loading && <WeeklyTotals slots={slots} oswald={oswald} employees={employees} />}

      {modal && <SlotModal modal={modal} onClose={() => setModal(null)} onSave={handleSaveSlot} onDelete={handleDeleteSlot} oswald={oswald} employees={employees} />}
      {viewSlot && <SlotViewModal slot={viewSlot} onClose={() => setViewSlot(null)} oswald={oswald} employees={employees} />}
    </div>
  );
}

function DayCard({ date, dayIndex, slots, isManager, onAddSlot, onEditSlot, currentEmployeeName, oswald, eventText, onEventChange, employees, markerFor, onToggleMarker, dayEvents }) {
  const today = sameDay(date, new Date());
  const [editingEvent, setEditingEvent] = useState(false);
  const [draft, setDraft] = useState(eventText);

  function handleBarClick(e, employee) {
    if (!isManager) return;
    if (markerFor(employee) === "cp") return;
    const rect = e.currentTarget.getBoundingClientRect();
    const x = e.clientX - rect.left;
    onAddSlot(dayIndex, employee, x, rect.width);
  }

  function commitEvent() {
    setEditingEvent(false);
    if (draft !== eventText) onEventChange(draft);
  }

  return (
    <div className="bg-white rounded-xl shadow-sm border border-slate-200 mb-2.5 overflow-hidden">
      <div className={"px-3 py-1.5 border-b flex items-center justify-between " + (today ? "bg-amber-50 border-amber-100" : "bg-emerald-50 border-emerald-100")}>
        <div className="font-bold text-emerald-900 text-sm" style={oswald}>{fmtDayLabel(date)}</div>
        {today && <span className="text-xs font-bold text-amber-600 uppercase">Aujourd'hui</span>}
      </div>

      <div className="flex">
        <div className="overflow-x-auto flex-1">
          <div style={{ minWidth: "700px" }}>
            <div className="relative h-4 ml-24 mr-2 mt-1.5">
              {HOURS.map((h) => (
                <div key={h} className="absolute text-slate-400 text-xs" style={{ left: pct(h * 60 - RANGE_START) + "%" }}>{h}h</div>
              ))}
            </div>
            <div className="px-2 pb-2">
              {employees.map((empObj) => {
                const emp = empObj.name;
                const empSlots = slots.filter((s) => s.employee === emp);
                const isMe = currentEmployeeName === emp;
                const colors = empColor(emp, employees);
                const marker = markerFor(emp);
                return (
                  <div key={emp} className="flex items-center mb-1">
                    <div className={"w-24 shrink-0 text-xs font-semibold pr-2 truncate " + (isMe ? "text-amber-600" : "text-slate-600")}>{emp}</div>
                    {isManager && (
                      <div className="flex gap-1 mr-2.5 shrink-0">
                        <button
                          onClick={() => onToggleMarker(emp, marker === "cp" ? null : "cp")}
                          className={"w-7 h-6 rounded text-xs font-bold " + (marker === "cp" ? "bg-slate-600 text-white" : "bg-slate-100 text-slate-400")}
                          title="Congé payé"
                        >
                          CP
                        </button>
                        <button
                          onClick={() => onToggleMarker(emp, marker === "cours" ? null : "cours")}
                          className={"w-7 h-6 rounded text-xs font-bold " + (marker === "cours" ? "bg-sky-600 text-white" : "bg-slate-100 text-slate-400")}
                          title="Cours"
                        >
                          C
                        </button>
                      </div>
                    )}
                    <div className="relative flex-1 h-6 bg-slate-50 rounded-md border border-slate-100" style={{ cursor: isManager && marker !== "cp" ? "pointer" : "default" }} onClick={(e) => handleBarClick(e, emp)}>
                      {HOURS.map((h) => (
                        <div key={h} className="absolute top-0 bottom-0 border-l border-slate-100" style={{ left: pct(h * 60 - RANGE_START) + "%" }} />
                      ))}
                      {marker === "cp" && (
                        <div className="absolute inset-0 rounded-md bg-slate-300 flex items-center justify-center text-slate-600 text-xs font-bold">
                          CP
                        </div>
                      )}
                      {marker === "cours" && (
                        <div
                          className="absolute inset-0 rounded-md flex items-center px-1.5"
                          style={{ background: "repeating-linear-gradient(45deg, rgba(2,132,199,0.12), rgba(2,132,199,0.12) 6px, transparent 6px, transparent 12px)" }}
                        >
                          <span className="bg-sky-100 text-sky-700 text-xs font-bold px-1 rounded">Cours</span>
                        </div>
                      )}
                      {empSlots.map((slot) => (
                        <div
                          key={slot.id}
                          onClick={(e) => { e.stopPropagation(); onEditSlot(slot); }}
                          className={"absolute top-0.5 bottom-0.5 rounded flex items-center px-1.5 text-white text-xs font-semibold overflow-hidden " + colors.bg}
                          style={{ left: pct(slot.start - RANGE_START) + "%", width: pct(slot.end - slot.start) + "%", cursor: "pointer" }}
                        >
                          <span className="truncate">{fmtTime(slot.start)}–{fmtTime(slot.end)}</span>
                          {slot.comment && <MessageSquare size={11} className="ml-1 shrink-0" />}
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>

        <div className="w-36 shrink-0 border-l border-slate-100 px-2 py-1.5 bg-slate-50">
          <div className="text-xs font-bold text-slate-400 uppercase tracking-wide mb-1">Événement</div>
          {dayEvents && dayEvents.length > 0 && (
            <div className="space-y-1 mb-1.5">
              {dayEvents.map((e) => (
                <div key={e.id} className={"text-xs rounded px-1.5 py-1 " + (e.kind === "anniversaire" ? "bg-violet-100 text-violet-700" : e.kind === "reservation" ? "bg-sky-100 text-sky-700" : e.kind === "tournoi" ? "bg-orange-100 text-orange-700" : "bg-amber-100 text-amber-700")}>
                  {summarizeEntry(e)}
                </div>
              ))}
            </div>
          )}
          {editingEvent ? (
            <textarea
              autoFocus
              value={draft}
              onChange={(e) => setDraft(e.target.value)}
              onBlur={commitEvent}
              rows={4}
              className="w-full text-xs border border-slate-200 rounded p-1"
            />
          ) : (
            <div
              onClick={() => isManager && setEditingEvent(true)}
              className={"text-xs text-slate-600 min-h-8 whitespace-pre-line " + (isManager ? "cursor-pointer" : "")}
            >
              {eventText || (isManager ? <span className="text-slate-300">Toucher pour ajouter</span> : <span className="text-slate-300">—</span>)}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function RecapView({ weekStart, slots, oswald }) {
  function countAt(dayIndex, hour) {
    const point = hour * 60;
    return slots.filter((s) => s.day === dayIndex && s.start <= point && s.end > point).length;
  }
  function shade(n) {
    if (n === 0) return "bg-slate-50 text-slate-300";
    if (n === 1) return "bg-emerald-50 text-emerald-700";
    if (n === 2) return "bg-emerald-200 text-emerald-800";
    if (n === 3) return "bg-emerald-400 text-emerald-950";
    return "bg-emerald-600 text-white";
  }
  const days = Array.from({ length: 7 }, (_, i) => addDays(weekStart, i));

  return (
    <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-x-auto">
      <div style={{ minWidth: "760px" }} className="p-2">
        <div className="flex mb-1">
          <div className="w-14 shrink-0" />
          {HOURS.map((h) => (
            <div key={h} className="flex-1 text-center text-xs text-slate-400">{h}h</div>
          ))}
        </div>
        {days.map((date, dayIndex) => (
          <div key={dayIndex} className="flex items-center mb-1">
            <div className="w-14 shrink-0 text-xs font-bold text-emerald-900" style={oswald}>{DAY_SHORT[dayIndex]} {date.getDate()}</div>
            {HOURS.map((h) => {
              const n = countAt(dayIndex, h);
              return (
                <div key={h} className={"flex-1 mx-0.5 h-7 rounded flex items-center justify-center text-xs font-bold " + shade(n)}>
                  {n > 0 ? n : ""}
                </div>
              );
            })}
          </div>
        ))}
      </div>
      <div className="px-3 pb-3 text-xs text-slate-400">Nombre d'employés présents par heure.</div>
    </div>
  );
}

function WeeklyTotals({ slots, oswald, employees }) {
  const totals = employees.map((empObj) => {
    const emp = empObj.name;
    const minutes = slots.filter((s) => s.employee === emp).reduce((sum, s) => sum + (s.end - s.start), 0);
    return { emp, hours: minutes / 60 };
  }).filter((t) => t.hours > 0);

  if (totals.length === 0) return null;

  function fmtHours(h) {
    const whole = Math.floor(h);
    const frac = h - whole;
    return frac === 0 ? `${whole}h` : `${whole}h${Math.round(frac * 60)}`;
  }

  return (
    <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-3 mt-3">
      <div className="text-xs font-bold text-slate-400 uppercase tracking-wide mb-2">Total d'heures cette semaine</div>
      <div className="flex flex-wrap gap-2">
        {totals.map(({ emp, hours }) => {
          const colors = empColor(emp, employees);
          return (
            <div key={emp} className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-slate-50">
              <span className={"w-2.5 h-2.5 rounded-full " + colors.bg} />
              <span className="text-sm font-semibold text-slate-700">{emp}</span>
              <span className="text-sm font-bold text-slate-900" style={oswald}>{fmtHours(hours)}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function SlotModal({ modal, onClose, onSave, onDelete, oswald, employees }) {
  const [start, setStart] = useState(modal.start);
  const [end, setEnd] = useState(modal.end);
  const [comment, setComment] = useState(modal.comment || "");
  const [error, setError] = useState("");
  const timeOptions = [];
  for (let m = RANGE_START; m <= RANGE_END; m += 30) timeOptions.push(m);
  const colors = empColor(modal.employee, employees);

  function submit() {
    if (end <= start) { setError("L'heure de fin doit être après l'heure de début."); return; }
    onSave({ start, end, comment });
  }

  function applyPreset(p) {
    if (modal.mode === "add") {
      onSave({ start: p.start, end: p.end, comment });
      return;
    }
    setStart(p.start);
    setEnd(p.end);
    setError("");
  }

  return (
    <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(15, 23, 42, 0.5)", zIndex: 50 }} onClick={onClose}>
      <div className="bg-white rounded-2xl w-full p-5 overflow-y-auto" style={{ maxWidth: "420px", maxHeight: "88vh" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-4">
          <div>
            <div className="text-xs text-slate-400 uppercase tracking-wide">{DAY_NAMES[modal.day]}</div>
            <div className={"font-bold " + colors.text} style={oswald}>{modal.employee}</div>
          </div>
          <button onClick={onClose} className="text-slate-400"><X size={20} /></button>
        </div>

        {modal.mode === "add" && (
          <div className="mb-3">
            <label className="text-xs font-semibold text-slate-500 mb-1 block">Commentaire (optionnel — visible par tous)</label>
            <textarea value={comment} onChange={(e) => setComment(e.target.value)} rows={2} placeholder="Ex: remplace Nathan, ouverture..." className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
          </div>
        )}

        <div className="grid grid-cols-2 gap-2 mb-4">
          {PRESETS.map((p) => (
            <button key={p.key} onClick={() => applyPreset(p)} className="text-left px-3 py-2 rounded-lg border border-slate-200 hover:bg-slate-50">
              <div className="text-sm font-semibold text-slate-700">{p.label}</div>
              <div className="text-xs text-slate-400">{fmtTime(p.start)} – {fmtTime(p.end)}</div>
            </button>
          ))}
        </div>
        {modal.mode === "add" && <div className="text-xs text-slate-400 -mt-2 mb-4">Un créneau pré-rempli s'enregistre directement, sans passer par "Enregistrer".</div>}

        <div className="grid grid-cols-2 gap-3 mb-3">
          <div>
            <label className="text-xs font-semibold text-slate-500 mb-1 block">Début</label>
            <select value={start} onChange={(e) => { setStart(Number(e.target.value)); setError(""); }} className="w-full border border-slate-200 rounded-lg px-2 py-2 text-sm">
              {timeOptions.map((m) => <option key={m} value={m}>{fmtTime(m)}</option>)}
            </select>
          </div>
          <div>
            <label className="text-xs font-semibold text-slate-500 mb-1 block">Fin</label>
            <select value={end} onChange={(e) => { setEnd(Number(e.target.value)); setError(""); }} className="w-full border border-slate-200 rounded-lg px-2 py-2 text-sm">
              {timeOptions.map((m) => <option key={m} value={m}>{fmtTime(m)}</option>)}
            </select>
          </div>
        </div>

        {error && <div className="text-rose-600 text-xs mb-3">{error}</div>}

        {modal.mode === "edit" && (
          <div className="mb-4">
            <label className="text-xs font-semibold text-slate-500 mb-1 block">Commentaire (visible par tous)</label>
            <textarea value={comment} onChange={(e) => setComment(e.target.value)} rows={2} placeholder="Ex: remplace Nathan, ouverture..." className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
          </div>
        )}

        <div className="flex gap-2">
          {modal.mode === "edit" && (
            <button onClick={onDelete} className="px-3 py-2 rounded-lg border border-rose-200 text-rose-600 flex items-center gap-1 text-sm font-semibold">
              <Trash2 size={16} /> Supprimer
            </button>
          )}
          <button onClick={submit} className="flex-1 py-2 rounded-lg bg-emerald-600 text-white font-bold text-sm">Enregistrer</button>
        </div>
      </div>
    </div>
  );
}

function SlotViewModal({ slot, onClose, oswald, employees }) {
  const colors = empColor(slot.employee, employees);
  return (
    <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(15, 23, 42, 0.5)", zIndex: 50 }} onClick={onClose}>
      <div className="bg-white rounded-2xl w-full p-5" style={{ maxWidth: "380px" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-3">
          <div className={"font-bold " + colors.text} style={oswald}>{slot.employee}</div>
          <button onClick={onClose} className="text-slate-400"><X size={20} /></button>
        </div>
        <div className="text-slate-700 font-semibold mb-2">{fmtTime(slot.start)} – {fmtTime(slot.end)}</div>
        {slot.comment ? (
          <div className="bg-slate-50 rounded-lg p-3 text-sm text-slate-600">{slot.comment}</div>
        ) : (
          <div className="text-sm text-slate-400">Aucun commentaire.</div>
        )}
      </div>
    </div>
  );
}

/* ---------------- ÉVÉNEMENT (JOURNALIER) ---------------- */

function kindStyle(kind) {
  if (kind === "anniversaire") return "bg-violet-50 text-violet-700";
  if (kind === "reservation") return "bg-sky-50 text-sky-700";
  if (kind === "tournoi") return "bg-orange-50 text-orange-700";
  return "bg-amber-50 text-amber-700";
}
function dotColor(kind) {
  if (kind === "anniversaire") return "bg-violet-500";
  if (kind === "reservation") return "bg-sky-500";
  if (kind === "tournoi") return "bg-orange-500";
  return "bg-amber-500";
}
function startOfMonth(d) {
  return new Date(d.getFullYear(), d.getMonth(), 1);
}
function endOfMonth(d) {
  return new Date(d.getFullYear(), d.getMonth() + 1, 0);
}

function EvenementTab({ profile, eventEntries, saveEventEntries, eventGroups, oswald }) {
  const [monthStart, setMonthStart] = useState(startOfMonth(new Date()));
  const [selectedDate, setSelectedDate] = useState(new Date());
  const [modal, setModal] = useState(null);
  const [search, setSearch] = useState("");

  function changeMonth(delta) {
    setMonthStart((m) => {
      const d = new Date(m);
      d.setMonth(d.getMonth() + delta);
      return d;
    });
  }
  function goToday() {
    const t = new Date();
    setMonthStart(startOfMonth(t));
    setSelectedDate(t);
  }

  async function handleSave(data) {
    let next;
    if (modal.mode === "add") {
      const id = "evt-" + Date.now();
      next = [...eventEntries, { id, date: modal.date, createdBy: profile.name, createdTs: Date.now(), ...data }];
    } else {
      next = eventEntries.map((e) => (e.id === modal.id ? { ...e, ...data } : e));
    }
    await saveEventEntries(next);
    setModal(null);
  }
  async function handleDelete() {
    await saveEventEntries(eventEntries.filter((e) => e.id !== modal.id));
    setModal(null);
  }

  const gridStart = mondayOf(startOfMonth(monthStart));
  const gridEndAnchor = endOfMonth(monthStart);
  const lastRowStart = mondayOf(gridEndAnchor);
  const gridEnd = addDays(lastRowStart, 6);

  const gridDays = [];
  let cursor = new Date(gridStart);
  while (cursor <= gridEnd) {
    gridDays.push(new Date(cursor));
    cursor = addDays(cursor, 1);
  }

  const allWeeks = [];
  for (let i = 0; i < gridDays.length; i += 7) {
    allWeeks.push(gridDays.slice(i, i + 7));
  }
  const todayISO = isoDate(new Date());

  const selectedISO = isoDate(selectedDate);
  const dayEntries = eventEntries.filter((e) => e.date === selectedISO);
  const monthLabel = monthStart.toLocaleDateString("fr-FR", { month: "long", year: "numeric" });

  const searchResults = search.trim()
    ? eventEntries
        .filter((e) => summarizeEntry(e).toLowerCase().includes(search.trim().toLowerCase()))
        .sort((a, b) => b.date.localeCompare(a.date))
    : null;

  return (
    <div>
      <div className="relative mb-3">
        <input
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          placeholder="Rechercher un évènement..."
          className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm"
        />
      </div>

      {searchResults ? (
        <div className="bg-white rounded-xl border border-slate-200 divide-y divide-slate-100">
          {searchResults.length === 0 ? (
            <div className="text-center text-slate-400 py-6 text-sm">Aucun résultat.</div>
          ) : (
            searchResults.map((e) => (
              <button
                key={e.id}
                onClick={() => setModal({ mode: "edit", ...e })}
                className="w-full text-left px-3 py-2.5 hover:bg-slate-50"
              >
                <div className="text-xs text-slate-400 mb-0.5 capitalize">{fmtDayLabel(new Date(e.date + "T00:00:00"))}</div>
                <div className="text-sm font-semibold text-slate-700">{summarizeEntry(e)}</div>
              </button>
            ))
          )}
        </div>
      ) : (
        <>
      <div className="flex items-center justify-between mb-3">
        <button onClick={() => changeMonth(-1)} className="p-2 rounded-lg bg-white shadow-sm border border-slate-200 text-slate-600">
          <ChevronLeft size={18} />
        </button>
        <div className="text-center">
          <div className="font-bold text-emerald-900 capitalize" style={oswald}>{monthLabel}</div>
          <button onClick={goToday} className="text-xs text-amber-600 font-semibold">Aujourd'hui</button>
        </div>
        <button onClick={() => changeMonth(1)} className="p-2 rounded-lg bg-white shadow-sm border border-slate-200 text-slate-600">
          <ChevronRight size={18} />
        </button>
      </div>

      <div className="grid grid-cols-7 gap-1 mb-1">
        {DAY_SHORT.map((d) => (
          <div key={d} className="text-center text-xs font-bold text-slate-400">{d}</div>
        ))}
      </div>

      <div className="grid grid-cols-7 gap-1 mb-4">
        {gridDays.map((date, i) => {
          const dISO = isoDate(date);
          const inMonth = date.getMonth() === monthStart.getMonth();
          const entries = eventEntries.filter((e) => e.date === dISO);
          const isSelected = dISO === selectedISO;
          const isToday = sameDay(date, new Date());
          return (
            <button
              key={i}
              onClick={() => setSelectedDate(date)}
              style={{ minHeight: "48px" }}
              className={
                "rounded-lg border p-1 flex flex-col items-start " +
                (isSelected ? "bg-emerald-600 border-emerald-600" : isToday ? "bg-amber-50 border-amber-200" : "bg-white border-slate-100") +
                (inMonth ? "" : " opacity-40")
              }
            >
              <span className={"text-xs font-bold " + (isSelected ? "text-white" : "text-slate-700")}>{date.getDate()}</span>
              <div className="flex flex-wrap gap-0.5 mt-auto">
                {entries.slice(0, 4).map((e) => (
                  <span key={e.id} className={"w-1.5 h-1.5 rounded-full " + dotColor(e.kind)} />
                ))}
              </div>
            </button>
          );
        })}
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
        <div className="px-3 py-2 bg-emerald-50 border-b border-emerald-100 font-bold text-emerald-900 text-sm" style={oswald}>
          {fmtDayLabel(selectedDate)}
        </div>
        <div className="p-3">
          {dayEntries.length === 0 ? (
            <div className="text-xs text-slate-400 mb-2">Aucun évènement.</div>
          ) : (
            <div className="space-y-1.5 mb-2">
              {dayEntries.map((e) => (
                <button
                  key={e.id}
                  onClick={() => setModal({ mode: "edit", ...e })}
                  className={"w-full text-left px-2.5 py-2 rounded-lg text-xs font-semibold " + kindStyle(e.kind)}
                >
                  {summarizeEntry(e)}
                </button>
              ))}
            </div>
          )}
          <button
            onClick={() => setModal({ mode: "add", id: null, date: selectedISO, kind: null })}
            className="text-xs font-semibold text-emerald-700 border border-dashed border-emerald-300 rounded-lg px-3 py-1.5 flex items-center gap-1"
          >
            <Plus size={14} /> Ajouter un évènement
          </button>
        </div>
      </div>

      <div className="mt-5">
        <div className="text-xs font-bold text-slate-400 uppercase tracking-wide mb-2">À venir</div>
        {(() => {
          const upcoming = eventEntries.filter((e) => e.date >= todayISO).sort((a, b) => a.date.localeCompare(b.date));
          if (upcoming.length === 0) {
            return <div className="text-center text-slate-400 py-6 text-sm">Aucun évènement à venir.</div>;
          }
          const byWeek = [];
          upcoming.forEach((e) => {
            const wk = isoDate(mondayOf(new Date(e.date + "T00:00:00")));
            let group = byWeek[byWeek.length - 1];
            if (!group || group.week !== wk) {
              group = { week: wk, items: [] };
              byWeek.push(group);
            }
            group.items.push(e);
          });
          return byWeek.map((group) => {
            const weekStartDate = new Date(group.week + "T00:00:00");
            const weekEndDate = addDays(weekStartDate, 6);
            return (
              <div key={group.week} className="bg-white rounded-xl border border-slate-200 mb-3 overflow-hidden">
                <div className="px-3 py-2 bg-slate-50 border-b border-slate-100 font-bold text-slate-700 text-sm">
                  Semaine du {fmtShort(weekStartDate)} au {fmtShort(weekEndDate)}
                </div>
                <div className="divide-y divide-slate-100">
                  {group.items.map((e) => (
                    <button
                      key={e.id}
                      onClick={() => setModal({ mode: "edit", ...e })}
                      className="w-full text-left px-3 py-2 flex items-start gap-2 hover:bg-slate-50"
                    >
                      <span className="text-xs font-bold text-slate-400 w-9 shrink-0">
                        {DAY_SHORT[(new Date(e.date + "T00:00:00").getDay() + 6) % 7]}
                      </span>
                      <span className="text-xs font-semibold text-slate-700 flex-1">{summarizeEntry(e)}</span>
                    </button>
                  ))}
                </div>
              </div>
            );
          });
        })()}
      </div>

        </>
      )}

      {modal && <EventModal modal={modal} eventGroups={eventGroups} onClose={() => setModal(null)} onSave={handleSave} onDelete={handleDelete} oswald={oswald} eventEntries={eventEntries} />}
    </div>
  );
}

function EventModal({ modal, eventGroups, onClose, onSave, onDelete, oswald, eventEntries }) {

  const [kind, setKind] = useState(modal.kind || null);

  const [slot, setSlot] = useState(modal.slot || "13h30");
  const [firstName, setFirstName] = useState(modal.firstName || "");
  const [lastName, setLastName] = useState(modal.lastName || "");
  const [age, setAge] = useState(modal.age || "");
  const [sportDuration, setSportDuration] = useState(modal.sportDuration || "1h");
  const [childCount, setChildCount] = useState(modal.childCount || "");
  const [phone, setPhone] = useState(modal.phone || "");
  const [email, setEmail] = useState(modal.email || "");
  const [depositPaid, setDepositPaid] = useState(modal.depositPaid || false);
  const [invitesSent, setInvitesSent] = useState(modal.invitesSent || false);
  const [fieldsBlocked, setFieldsBlocked] = useState(modal.fieldsBlocked || false);

  const [forWhom, setForWhom] = useState(modal.forWhom || "");
  const [isBar, setIsBar] = useState(modal.isBar || false);
  const [isSport, setIsSport] = useState(modal.isSport || false);
  const [time, setTime] = useState(modal.time || "");
  const [description, setDescription] = useState(modal.description || "");
  const [confirmed, setConfirmed] = useState(modal.confirmed || false);

  const [subtype, setSubtype] = useState(modal.subtype || null);
  const [level, setLevel] = useState(modal.level || FFT_LEVELS[0]);
  const [loisirsType, setLoisirsType] = useState(modal.loisirsType || LOISIRS_TYPES[0]);

  const [category, setCategory] = useState(modal.category || eventGroups[0]?.name || "");
  const [comment, setComment] = useState(modal.comment || "");

  const [error, setError] = useState("");

  const takenSlots = eventEntries
    .filter((e) => e.kind === "anniversaire" && e.date === modal.date && e.id !== modal.id)
    .map((e) => e.slot);

  function submit() {
    if (kind === "anniversaire") {
      if (!firstName.trim()) { setError("Le prénom est requis."); return; }
      if (takenSlots.includes(slot)) { setError("Ce créneau est déjà réservé pour un autre anniversaire ce jour-là."); return; }
      onSave({ kind, slot, firstName: firstName.trim(), lastName: lastName.trim(), age, sportDuration, childCount, phone, email, depositPaid, invitesSent, fieldsBlocked });
    } else if (kind === "reservation") {
      if (!forWhom.trim()) { setError("Merci de préciser pour qui est la réservation."); return; }
      onSave({ kind, forWhom: forWhom.trim(), isBar, isSport, time, description, confirmed });
    } else if (kind === "tournoi") {
      if (subtype === "fft") {
        onSave({ kind, subtype, level, comment: comment.trim() });
      } else {
        if (loisirsType === "Autres" && !comment.trim()) { setError("Merci de préciser le type de tournoi."); return; }
        onSave({ kind, subtype, loisirsType, comment: loisirsType === "Autres" ? comment.trim() : "" });
      }
    } else {
      if (!comment.trim()) { setError("Merci d'ajouter un commentaire."); return; }
      onSave({ kind: "simple", category, comment: comment.trim() });
    }
  }

  if (modal.mode === "add" && !kind) {
    return (
      <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(15, 23, 42, 0.5)", zIndex: 50 }} onClick={onClose}>
        <div className="bg-white rounded-2xl w-full p-5" style={{ maxWidth: "380px" }} onClick={(e) => e.stopPropagation()}>
          <div className="flex items-center justify-between mb-4">
            <div className="font-bold text-emerald-900" style={oswald}>Type d'évènement</div>
            <button onClick={onClose} className="text-slate-400"><X size={20} /></button>
          </div>
          <div className="grid grid-cols-2 gap-2">
            <button onClick={() => setKind("anniversaire")} className="py-3 rounded-lg border border-slate-200 text-sm font-semibold text-slate-700 hover:bg-slate-50">
              Anniversaire
            </button>
            <button onClick={() => setKind("reservation")} className="py-3 rounded-lg border border-slate-200 text-sm font-semibold text-slate-700 hover:bg-slate-50">
              Réservation
            </button>
            <button onClick={() => setKind("tournoi")} className="py-3 rounded-lg border border-slate-200 text-sm font-semibold text-slate-700 hover:bg-slate-50">
              Tournoi
            </button>
            {eventGroups.map((g) => (
              <button key={g.id} onClick={() => { setKind("simple"); setCategory(g.name); }} className="py-3 rounded-lg border border-slate-200 text-sm font-semibold text-slate-700 hover:bg-slate-50">
                {g.name}
              </button>
            ))}
          </div>
        </div>
      </div>
    );
  }

  if (modal.mode === "add" && kind === "tournoi" && !subtype) {
    return (
      <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(15, 23, 42, 0.5)", zIndex: 50 }} onClick={onClose}>
        <div className="bg-white rounded-2xl w-full p-5" style={{ maxWidth: "380px" }} onClick={(e) => e.stopPropagation()}>
          <div className="flex items-center justify-between mb-4">
            <div className="font-bold text-emerald-900" style={oswald}>Type de tournoi</div>
            <button onClick={onClose} className="text-slate-400"><X size={20} /></button>
          </div>
          <div className="grid grid-cols-1 gap-2">
            <button onClick={() => setSubtype("fft")} className="py-3 rounded-lg border border-slate-200 text-sm font-semibold text-slate-700 hover:bg-slate-50">
              Tournoi FFT
            </button>
            <button onClick={() => setSubtype("loisirs")} className="py-3 rounded-lg border border-slate-200 text-sm font-semibold text-slate-700 hover:bg-slate-50">
              Tournoi loisirs
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(15, 23, 42, 0.5)", zIndex: 50 }} onClick={onClose}>
      <div className="bg-white rounded-2xl w-full p-5 overflow-y-auto" style={{ maxWidth: "420px", maxHeight: "88vh" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-4">
          <div className="font-bold text-emerald-900" style={oswald}>
            {kind === "anniversaire" ? "Anniversaire" : kind === "reservation" ? "Réservation" : kind === "tournoi" ? (subtype === "fft" ? "Tournoi FFT" : "Tournoi loisirs") : category || "Évènement"}
          </div>
          <button onClick={onClose} className="text-slate-400"><X size={20} /></button>
        </div>

        {kind === "anniversaire" && (
          <>
            <label className="text-xs font-semibold text-slate-500 mb-1 block">Créneau</label>
            <div className="grid grid-cols-3 gap-1.5 mb-3">
              {ANNIV_SLOTS.map((s) => {
                const taken = takenSlots.includes(s);
                return (
                  <button
                    key={s}
                    onClick={() => !taken && setSlot(s)}
                    disabled={taken}
                    className={
                      "py-2 rounded-lg text-xs font-bold border " +
                      (taken ? "border-slate-100 text-slate-300 bg-slate-50 line-through cursor-not-allowed" : slot === s ? "bg-emerald-600 text-white border-emerald-600" : "border-slate-200 text-slate-500")
                    }
                  >
                    {s}
                  </button>
                );
              })}
            </div>
            <div className="grid grid-cols-2 gap-3 mb-3">
              <div>
                <label className="text-xs font-semibold text-slate-500 mb-1 block">Prénom</label>
                <input value={firstName} onChange={(e) => setFirstName(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
              </div>
              <div>
                <label className="text-xs font-semibold text-slate-500 mb-1 block">Nom</label>
                <input value={lastName} onChange={(e) => setLastName(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
              </div>
            </div>
            <div className="grid grid-cols-2 gap-3 mb-3">
              <div>
                <label className="text-xs font-semibold text-slate-500 mb-1 block">Âge</label>
                <input type="number" value={age} onChange={(e) => setAge(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
              </div>
              <div>
                <label className="text-xs font-semibold text-slate-500 mb-1 block">Nombre d'enfants</label>
                <input type="number" value={childCount} onChange={(e) => setChildCount(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
              </div>
            </div>
            <label className="text-xs font-semibold text-slate-500 mb-1 block">Durée du sport</label>
            <div className="grid grid-cols-2 gap-1.5 mb-3">
              {ANNIV_DURATIONS.map((d) => (
                <button key={d} onClick={() => setSportDuration(d)} className={"py-2 rounded-lg text-xs font-bold border " + (sportDuration === d ? "bg-emerald-600 text-white border-emerald-600" : "border-slate-200 text-slate-500")}>
                  {d}
                </button>
              ))}
            </div>
            <div className="grid grid-cols-2 gap-3 mb-3">
              <div>
                <label className="text-xs font-semibold text-slate-500 mb-1 block">Téléphone</label>
                <input value={phone} onChange={(e) => setPhone(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
              </div>
              <div>
                <label className="text-xs font-semibold text-slate-500 mb-1 block">Email</label>
                <input value={email} onChange={(e) => setEmail(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
              </div>
            </div>
            <div className="space-y-2 mb-3">
              <label className="flex items-center gap-2 text-sm text-slate-700">
                <input type="checkbox" checked={depositPaid} onChange={(e) => setDepositPaid(e.target.checked)} /> Acompte payé
              </label>
              <label className="flex items-center gap-2 text-sm text-slate-700">
                <input type="checkbox" checked={invitesSent} onChange={(e) => setInvitesSent(e.target.checked)} /> Cartons envoyés
              </label>
              <label className="flex items-center gap-2 text-sm text-slate-700">
                <input type="checkbox" checked={fieldsBlocked} onChange={(e) => setFieldsBlocked(e.target.checked)} /> Terrains bloqués
              </label>
            </div>
          </>
        )}

        {kind === "reservation" && (
          <>
            <label className="text-xs font-semibold text-slate-500 mb-1 block">Pour qui</label>
            <input value={forWhom} onChange={(e) => setForWhom(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" placeholder="Ex: Entreprise XX" />
            <div className="flex gap-4 mb-3">
              <label className="flex items-center gap-2 text-sm text-slate-700">
                <input type="checkbox" checked={isBar} onChange={(e) => setIsBar(e.target.checked)} /> Bar
              </label>
              <label className="flex items-center gap-2 text-sm text-slate-700">
                <input type="checkbox" checked={isSport} onChange={(e) => setIsSport(e.target.checked)} /> Sport
              </label>
            </div>
            <label className="text-xs font-semibold text-slate-500 mb-1 block">Heure (pour le récap, optionnel)</label>
            <input value={time} onChange={(e) => setTime(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" placeholder="Ex: 18h" />
            <label className="text-xs font-semibold text-slate-500 mb-1 block">Description (nombre de personnes, horaire...)</label>
            <textarea value={description} onChange={(e) => setDescription(e.target.value)} rows={2} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" />
            <label className="flex items-center gap-2 text-sm text-slate-700 mb-3">
              <input type="checkbox" checked={confirmed} onChange={(e) => setConfirmed(e.target.checked)} /> Confirmé
            </label>
          </>
        )}

        {kind === "tournoi" && subtype === "fft" && (
          <>
            <label className="text-xs font-semibold text-slate-500 mb-1 block">Niveau Tournoi FFT</label>
            <div className="grid grid-cols-3 gap-1.5 mb-3">
              {FFT_LEVELS.map((l) => (
                <button key={l} onClick={() => setLevel(l)} className={"py-2 rounded-lg text-xs font-bold border " + (level === l ? "bg-orange-500 text-white border-orange-500" : "border-slate-200 text-slate-500")}>
                  {l}
                </button>
              ))}
            </div>
            <label className="text-xs font-semibold text-slate-500 mb-1 block">Commentaire (optionnel)</label>
            <textarea value={comment} onChange={(e) => setComment(e.target.value)} rows={2} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" />
          </>
        )}

        {kind === "tournoi" && subtype === "loisirs" && (
          <>
            <label className="text-xs font-semibold text-slate-500 mb-1 block">Type de tournoi loisirs</label>
            <div className="grid grid-cols-1 gap-1.5 mb-3">
              {LOISIRS_TYPES.map((t) => (
                <button key={t} onClick={() => setLoisirsType(t)} className={"py-2 rounded-lg text-xs font-bold border text-left px-3 " + (loisirsType === t ? "bg-orange-500 text-white border-orange-500" : "border-slate-200 text-slate-500")}>
                  {t}
                </button>
              ))}
            </div>
            {loisirsType === "Autres" && (
              <>
                <label className="text-xs font-semibold text-slate-500 mb-1 block">Précisez</label>
                <textarea value={comment} onChange={(e) => setComment(e.target.value)} rows={2} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" />
              </>
            )}
          </>
        )}

        {kind === "simple" && (
          <>
            <label className="text-xs font-semibold text-slate-500 mb-1 block">Groupe</label>
            <select value={category} onChange={(e) => setCategory(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3">
              {eventGroups.map((g) => <option key={g.id} value={g.name}>{g.name}</option>)}
            </select>
            <label className="text-xs font-semibold text-slate-500 mb-1 block">Commentaire</label>
            <textarea value={comment} onChange={(e) => setComment(e.target.value)} rows={3} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" />
          </>
        )}

        {error && <div className="text-rose-600 text-xs mb-3">{error}</div>}

        <div className="flex gap-2">
          {modal.mode === "edit" && (
            <button onClick={onDelete} className="px-3 py-2 rounded-lg border border-rose-200 text-rose-600 flex items-center gap-1 text-sm font-semibold">
              <Trash2 size={16} /> Supprimer
            </button>
          )}
          <button onClick={submit} className="flex-1 py-2 rounded-lg bg-emerald-600 text-white font-bold text-sm">Enregistrer</button>
        </div>
      </div>
    </div>
  );
}

/* ---------------- CHECKLISTS (OUVERTURE / FERMETURE) ---------------- */


function ChecklistTab({ type, title, seed, isManager, profile, oswald }) {
  const [tasks, setTasks] = useState([]);
  const [done, setDone] = useState({});
  const [finished, setFinished] = useState(null);
  const [loading, setLoading] = useState(true);
  const [manageOpen, setManageOpen] = useState(false);
  const [taskModal, setTaskModal] = useState(null);
  const [valueModal, setValueModal] = useState(null);

  const todayISO = isoDate(new Date());
  const doneKey = `checklist:${type}:${todayISO}`;
  const finishedKey = `process-done:${type}:${todayISO}`;

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    (async () => {
      let t = await storageGet(`tasks:${type}:v4`, null);
      if (t === null) {
        t = seed;
        await storageSet(`tasks:${type}:v4`, t);
      }
      const d = await storageGet(doneKey, {});
      const f = await storageGet(finishedKey, null);
      if (!cancelled) {
        setTasks(t);
        setDone(d);
        setFinished(f);
        setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [type, doneKey, finishedKey]);

  async function toggleTask(task) {
    if (done[task.id]) {
      const next = { ...done };
      delete next[task.id];
      setDone(next);
      await storageSet(doneKey, next);
      return;
    }
    if (task.valueLabel) {
      setValueModal({ taskId: task.id, label: task.valueLabel });
      return;
    }
    const next = { ...done, [task.id]: { by: profile.name, ts: Date.now() } };
    setDone(next);
    await storageSet(doneKey, next);
  }

  async function confirmValue(value) {
    const next = { ...done, [valueModal.taskId]: { by: profile.name, ts: Date.now(), value } };
    setDone(next);
    await storageSet(doneKey, next);
    setValueModal(null);
  }

  async function finishProcess() {
    const entry = { by: profile.name, ts: Date.now() };
    setFinished(entry);
    await storageSet(finishedKey, entry);
  }

  async function saveTask(data) {
    let next;
    if (taskModal.mode === "add") {
      const id = type[0] + Date.now();
      next = [...tasks, { id, group: data.group, text: data.text, freq: data.freq, valueLabel: data.valueLabel }];
    } else {
      next = tasks.map((t) => (t.id === taskModal.id ? { ...t, group: data.group, text: data.text, freq: data.freq, valueLabel: data.valueLabel } : t));
    }
    setTasks(next);
    await storageSet(`tasks:${type}:v4`, next);
    setTaskModal(null);
  }
  async function deleteTask() {
    const next = tasks.filter((t) => t.id !== taskModal.id);
    setTasks(next);
    await storageSet(`tasks:${type}:v4`, next);
    setTaskModal(null);
  }

  const groups = useMemo(() => {
    const order = [];
    const map = {};
    tasks.forEach((t) => {
      if (!map[t.group]) { map[t.group] = []; order.push(t.group); }
      map[t.group].push(t);
    });
    return order.map((g) => ({ name: g, items: map[g] }));
  }, [tasks]);

  const doneCount = Object.keys(done).length;

  if (loading) return <div className="text-center text-slate-400 py-10">Chargement…</div>;

  return (
    <div>
      <div className="flex items-center justify-between mb-3">
        <div>
          <div className="font-bold text-emerald-900" style={oswald}>{title}</div>
          <div className="text-xs text-slate-400">{doneCount}/{tasks.length} tâches faites aujourd'hui</div>
        </div>
        {isManager && (
          <button onClick={() => setManageOpen((o) => !o)} className="text-xs font-semibold text-emerald-700 border border-emerald-200 rounded-lg px-3 py-1.5">
            {manageOpen ? "Fermer la gestion" : "Gérer les tâches"}
          </button>
        )}
      </div>

      {finished && (
        <div className="bg-emerald-50 border border-emerald-100 rounded-lg px-3 py-2 mb-3 text-sm text-emerald-800">
          Process terminé par <span className="font-semibold">{finished.by}</span> à {new Date(finished.ts).toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })}
        </div>
      )}

      {manageOpen && isManager && (
        <button onClick={() => setTaskModal({ mode: "add", group: groups[0]?.name || "", text: "", freq: "", valueLabel: "" })} className="w-full mb-3 py-2 rounded-lg border border-dashed border-emerald-300 text-emerald-700 text-sm font-semibold flex items-center justify-center gap-1">
          <Plus size={16} /> Ajouter une tâche
        </button>
      )}

      {groups.map((g) => (
        <div key={g.name} className="bg-white rounded-xl shadow-sm border border-slate-200 mb-3 overflow-hidden">
          <div className="px-3 py-2 bg-slate-50 border-b border-slate-100 font-bold text-slate-700 text-sm">{g.name}</div>
          <div className="divide-y divide-slate-100">
            {g.items.map((task) => {
              const d = done[task.id];
              return (
                <div key={task.id} className="flex items-start gap-2 px-3 py-2">
                  <button onClick={() => toggleTask(task)} className={"mt-0.5 w-5 h-5 rounded border shrink-0 flex items-center justify-center " + (d ? "bg-emerald-600 border-emerald-600" : "border-slate-300")}>
                    {d && <Check size={14} className="text-white" />}
                  </button>
                  <div className="flex-1 min-w-0">
                    <div className={"text-sm " + (d ? "text-slate-400 line-through" : "text-slate-700")}>{task.text}</div>
                    <div className="text-xs text-slate-400">
                      {[task.freq, d ? `${d.value !== undefined ? d.value + " · " : ""}fait par ${d.by} à ${new Date(d.ts).toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })}` : ""].filter(Boolean).join(" · ")}
                    </div>
                  </div>
                  {isManager && manageOpen && (
                    <button onClick={() => setTaskModal({ mode: "edit", id: task.id, group: task.group, text: task.text, freq: task.freq, valueLabel: task.valueLabel || "" })} className="text-slate-400 shrink-0">
                      <Pencil size={14} />
                    </button>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      ))}

      <button onClick={finishProcess} className="w-full py-3 rounded-xl bg-amber-400 text-emerald-900 font-bold mt-2">
        Terminer le process {type === "ouverture" ? "d'ouverture" : "de fermeture"}
      </button>

      {taskModal && <TaskModal modal={taskModal} groups={groups.map((g) => g.name)} onClose={() => setTaskModal(null)} onSave={saveTask} onDelete={deleteTask} oswald={oswald} />}
      {valueModal && <ValueModal modal={valueModal} onClose={() => setValueModal(null)} onConfirm={confirmValue} oswald={oswald} />}
    </div>
  );
}

function ValueModal({ modal, onClose, onConfirm, oswald }) {
  const [value, setValue] = useState("");
  const [error, setError] = useState("");

  function submit() {
    if (!value.trim()) { setError("Merci de renseigner une valeur."); return; }
    onConfirm(value.trim());
  }

  return (
    <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(15, 23, 42, 0.5)", zIndex: 50 }} onClick={onClose}>
      <div className="bg-white rounded-2xl w-full p-5" style={{ maxWidth: "340px" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-3">
          <div className="font-bold text-emerald-900" style={oswald}>{modal.label}</div>
          <button onClick={onClose} className="text-slate-400"><X size={20} /></button>
        </div>
        <input
          autoFocus
          value={value}
          onChange={(e) => { setValue(e.target.value); setError(""); }}
          onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
          className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-2"
          placeholder="Ex: 4°C"
        />
        {error && <div className="text-rose-600 text-xs mb-2">{error}</div>}
        <button onClick={submit} className="w-full py-2 rounded-lg bg-emerald-600 text-white font-bold text-sm">Valider la tâche</button>
      </div>
    </div>
  );
}

function TaskModal({ modal, groups, onClose, onSave, onDelete, oswald }) {
  const [group, setGroup] = useState(modal.group);
  const [text, setText] = useState(modal.text);
  const [freq, setFreq] = useState(modal.freq);
  const [needsValue, setNeedsValue] = useState(!!modal.valueLabel);
  const [valueLabel, setValueLabel] = useState(modal.valueLabel || "");
  const [error, setError] = useState("");

  function submit() {
    if (!text.trim() || !group.trim()) { setError("Le groupe et le texte de la tâche sont requis."); return; }
    if (needsValue && !valueLabel.trim()) { setError("Précisez ce qui doit être noté (ex: Température)."); return; }
    onSave({ group: group.trim(), text: text.trim(), freq: freq.trim(), valueLabel: needsValue ? valueLabel.trim() : "" });
  }

  return (
    <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(15, 23, 42, 0.5)", zIndex: 50 }} onClick={onClose}>
      <div className="bg-white rounded-2xl w-full p-5" style={{ maxWidth: "420px" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-4">
          <div className="font-bold text-emerald-900" style={oswald}>{modal.mode === "add" ? "Nouvelle tâche" : "Modifier la tâche"}</div>
          <button onClick={onClose} className="text-slate-400"><X size={20} /></button>
        </div>

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Groupe</label>
        <input list="groups-list" value={group} onChange={(e) => setGroup(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" placeholder="Ex: Bar & accueil" />
        <datalist id="groups-list">
          {groups.map((g) => <option key={g} value={g} />)}
        </datalist>

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Tâche</label>
        <textarea value={text} onChange={(e) => setText(e.target.value)} rows={2} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" placeholder="Ex: Nettoyer la tireuse" />

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Fréquence (optionnel)</label>
        <input value={freq} onChange={(e) => setFreq(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" placeholder="Ex: 1er lundi du mois" />

        <label className="flex items-center gap-2 text-xs font-semibold text-slate-500 mb-2">
          <input type="checkbox" checked={needsValue} onChange={(e) => setNeedsValue(e.target.checked)} />
          Demander une valeur à la validation (ex: température)
        </label>
        {needsValue && (
          <input value={valueLabel} onChange={(e) => setValueLabel(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" placeholder="Ex: Température (°C)" />
        )}

        {error && <div className="text-rose-600 text-xs mb-3">{error}</div>}

        <div className="flex gap-2">
          {modal.mode === "edit" && (
            <button onClick={onDelete} className="px-3 py-2 rounded-lg border border-rose-200 text-rose-600 flex items-center gap-1 text-sm font-semibold">
              <Trash2 size={16} /> Supprimer
            </button>
          )}
          <button onClick={submit} className="flex-1 py-2 rounded-lg bg-emerald-600 text-white font-bold text-sm">Enregistrer</button>
        </div>
      </div>
    </div>
  );
}

/* ---------------- ÉVÉNEMENT (PLAN DE SALLE) ---------------- */

/* ---------------- TO DO (ACTIONS + PROJETS) ---------------- */

function TodoTab({ profile, isManager, actions, saveActions, managerActions, saveManagerActions, projects, saveProjects, employees, oswald }) {
  const [personFilter, setPersonFilter] = useState(profile.role === "employe" ? profile.name : "all");
  const allNames = [...MANAGERS, ...employees.map((e) => e.name)];

  useEffect(() => {
    setPersonFilter(profile.role === "employe" ? profile.name : "all");
  }, [profile.name, profile.role]);

  return (
    <div>
      <div className="flex items-center gap-2 mb-4">
        <span className="text-xs font-semibold text-slate-500 shrink-0">Filtrer par personne</span>
        <select value={personFilter} onChange={(e) => setPersonFilter(e.target.value)} className="border border-slate-200 rounded-lg px-2 py-1.5 text-sm flex-1">
          <option value="all">Tout le monde</option>
          {allNames.map((n) => (
            <option key={n} value={n}>{n}</option>
          ))}
        </select>
      </div>

      {isManager && (
        <ActionsSection
          profile={profile}
          actions={managerActions}
          saveActions={saveManagerActions}
          allNames={MANAGERS}
          personFilter={personFilter}
          oswald={oswald}
          title="Actions Managers"
        />
      )}

      <ActionsSection profile={profile} actions={actions} saveActions={saveActions} allNames={allNames} personFilter={personFilter} oswald={oswald} />
      <ProjectsSection profile={profile} projects={projects} saveProjects={saveProjects} allNames={allNames} personFilter={personFilter} oswald={oswald} />
    </div>
  );
}

function ActionsSection({ profile, actions, saveActions, allNames, personFilter, oswald, title }) {
  const [modal, setModal] = useState(null);
  const [showArchive, setShowArchive] = useState(false);

  const filtered = personFilter === "all" ? actions : actions.filter((a) => hasAssignee(a, personFilter));

  const todo = filtered
    .filter((a) => !a.done)
    .sort((a, b) => {
      const pa = PRIORITY_RANK[a.priority] ?? 3;
      const pb = PRIORITY_RANK[b.priority] ?? 3;
      if (pa !== pb) return pa - pb;
      if (a.date && b.date) return a.date.localeCompare(b.date);
      if (a.date) return -1;
      if (b.date) return 1;
      return 0;
    });
  const done = filtered.filter((a) => a.done);

  async function toggleDone(action) {
    const next = actions.map((a) =>
      a.id === action.id
        ? a.done
          ? { ...a, done: false, doneBy: null, doneTs: null }
          : { ...a, done: true, doneBy: profile.name, doneTs: Date.now() }
        : a
    );
    await saveActions(next);
  }

  async function handleSave(data) {
    let next;
    if (modal.mode === "add") {
      const id = "act-" + Date.now();
      next = [...actions, { id, text: data.text, date: data.date, assignees: data.assignees, priority: data.priority, comment: data.comment, freq: "", done: false, doneBy: null, doneTs: null, createdBy: profile.name, createdTs: Date.now() }];
    } else {
      next = actions.map((a) => (a.id === modal.id ? { ...a, text: data.text, date: data.date, assignees: data.assignees, priority: data.priority, comment: data.comment } : a));
    }
    await saveActions(next);
    setModal(null);
  }
  async function handleDelete() {
    await saveActions(actions.filter((a) => a.id !== modal.id));
    setModal(null);
  }

  return (
    <div className="mb-6">
      <div className="flex items-center justify-between mb-3">
        <div className="font-bold text-emerald-900" style={oswald}>{title || "Actions"}</div>
        <button
          onClick={() => setModal({ mode: "add", id: null, text: "", date: "", assignees: ["Tout le monde"], priority: "moyenne", comment: "" })}
          className="text-xs font-semibold text-white bg-emerald-600 rounded-lg px-3 py-1.5 flex items-center gap-1"
        >
          <Plus size={14} /> Nouvelle action
        </button>
      </div>

      {todo.length === 0 ? (
        <div className="text-center text-slate-400 py-6 text-sm">Aucune action en attente.</div>
      ) : (
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 divide-y divide-slate-100 mb-3">
          {todo.map((a) => (
            <ActionRow key={a.id} action={a} onToggle={() => toggleDone(a)} onEdit={() => setModal({ mode: "edit", ...a })} oswald={oswald} />
          ))}
        </div>
      )}

      {done.length > 0 && (
        <div>
          <button onClick={() => setShowArchive((s) => !s)} className="text-xs font-semibold text-slate-400 mb-2">
            {showArchive ? "Masquer" : "Voir"} les archives ({done.length})
          </button>
          {showArchive && (
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 divide-y divide-slate-100">
              {done.map((a) => (
                <ActionRow key={a.id} action={a} onToggle={() => toggleDone(a)} onEdit={() => setModal({ mode: "edit", ...a })} oswald={oswald} />
              ))}
            </div>
          )}
        </div>
      )}

      {modal && <ActionModal modal={modal} onClose={() => setModal(null)} onSave={handleSave} onDelete={handleDelete} oswald={oswald} allNames={allNames} />}
    </div>
  );
}

function ActionRow({ action, onToggle, onEdit, oswald }) {
  const overdue = action.date && !action.done && action.date < isoDate(new Date());
  const prio = PRIORITIES.find((p) => p.key === action.priority);
  return (
    <div className={"flex items-start gap-2 px-3 py-2.5 " + (overdue ? "bg-rose-50" : "")}>
      <button onClick={onToggle} className={"mt-0.5 w-5 h-5 rounded border shrink-0 flex items-center justify-center " + (action.done ? "bg-emerald-600 border-emerald-600" : "border-slate-300")}>
        {action.done && <Check size={14} className="text-white" />}
      </button>
      <div className="flex-1 min-w-0">
        <div className={"text-sm " + (action.done ? "text-slate-400 line-through" : overdue ? "text-rose-700 font-semibold" : "text-slate-700")}>{action.text}</div>
        <div className="flex flex-wrap gap-1.5 mt-1">
          {overdue && <span className="text-xs bg-rose-600 text-white font-bold rounded px-1.5 py-0.5">En retard</span>}
          {prio && <span className={"text-xs font-bold rounded px-1.5 py-0.5 " + prio.bg + " " + prio.text}>{prio.label}</span>}
          {action.date && <span className={"text-xs rounded px-1.5 py-0.5 " + (overdue ? "bg-rose-100 text-rose-700" : "bg-slate-100 text-slate-500")}>{new Date(action.date).toLocaleDateString("fr-FR", { day: "numeric", month: "short" })}</span>}
          {(action.assignees || (action.assignee ? [action.assignee] : [])).map((n) => (
            <span key={n} className="text-xs bg-emerald-50 text-emerald-700 rounded px-1.5 py-0.5">{n}</span>
          ))}
          {action.freq && <span className="text-xs bg-amber-50 text-amber-700 rounded px-1.5 py-0.5">{action.freq}</span>}
          {action.done && <span className="text-xs text-slate-400">fait par {action.doneBy} le {new Date(action.doneTs).toLocaleDateString("fr-FR")}</span>}
        </div>
        {action.comment && <div className="text-xs text-slate-500 mt-1">{action.comment}</div>}
      </div>
      <button onClick={onEdit} className="text-slate-400 shrink-0"><Pencil size={14} /></button>
    </div>
  );
}

function ActionModal({ modal, onClose, onSave, onDelete, oswald, allNames }) {
  const [text, setText] = useState(modal.text);
  const [date, setDate] = useState(modal.date || "");
  const [assignees, setAssignees] = useState(modal.assignees || (modal.assignee ? [modal.assignee] : ["Tout le monde"]));
  const [priority, setPriority] = useState(modal.priority || "moyenne");
  const [comment, setComment] = useState(modal.comment || "");
  const [error, setError] = useState("");

  function toggleAssignee(n) {
    setAssignees((prev) => (prev.includes(n) ? prev.filter((x) => x !== n) : [...prev, n]));
  }

  function submit() {
    if (!text.trim()) { setError("Le texte de l'action est requis."); return; }
    if (assignees.length === 0) { setError("Sélectionnez au moins une personne."); return; }
    onSave({ text: text.trim(), date, assignees, priority, comment });
  }

  return (
    <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(15, 23, 42, 0.5)", zIndex: 50 }} onClick={onClose}>
      <div className="bg-white rounded-2xl w-full p-5 overflow-y-auto" style={{ maxWidth: "420px", maxHeight: "88vh" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-4">
          <div className="font-bold text-emerald-900" style={oswald}>{modal.mode === "add" ? "Nouvelle action" : "Modifier l'action"}</div>
          <button onClick={onClose} className="text-slate-400"><X size={20} /></button>
        </div>

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Action</label>
        <textarea value={text} onChange={(e) => setText(e.target.value)} rows={2} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" placeholder="Ex: Recommander des verres" />

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Date</label>
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="w-full border border-slate-200 rounded-lg px-2 py-2 text-sm mb-3" />

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Assigné à (plusieurs possibles)</label>
        <div className="flex flex-wrap gap-1.5 mb-3">
          <button onClick={() => toggleAssignee("Tout le monde")} className={"text-xs font-semibold rounded-full px-2.5 py-1 border " + (assignees.includes("Tout le monde") ? "bg-emerald-600 text-white border-emerald-600" : "border-slate-200 text-slate-500")}>
            Tout le monde
          </button>
          {allNames.map((n) => (
            <button key={n} onClick={() => toggleAssignee(n)} className={"text-xs font-semibold rounded-full px-2.5 py-1 border " + (assignees.includes(n) ? "bg-emerald-600 text-white border-emerald-600" : "border-slate-200 text-slate-500")}>
              {n}
            </button>
          ))}
        </div>

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Priorité</label>
        <div className="grid grid-cols-3 gap-1.5 mb-3">
          {PRIORITIES.map((p) => (
            <button key={p.key} onClick={() => setPriority(p.key)} className={"py-2 rounded-lg text-xs font-bold border " + (priority === p.key ? p.bg + " " + p.text + " border-transparent" : "border-slate-200 text-slate-400")}>
              {p.label}
            </button>
          ))}
        </div>

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Commentaire</label>
        <textarea value={comment} onChange={(e) => setComment(e.target.value)} rows={2} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" />

        {error && <div className="text-rose-600 text-xs mb-3">{error}</div>}

        <div className="flex gap-2">
          {modal.mode === "edit" && (
            <button onClick={onDelete} className="px-3 py-2 rounded-lg border border-rose-200 text-rose-600 flex items-center gap-1 text-sm font-semibold">
              <Trash2 size={16} /> Supprimer
            </button>
          )}
          <button onClick={submit} className="flex-1 py-2 rounded-lg bg-emerald-600 text-white font-bold text-sm">Enregistrer</button>
        </div>
      </div>
    </div>
  );
}

const PROJECT_STATUSES = [
  { key: "a_faire", label: "À faire", bg: "bg-slate-100", text: "text-slate-600" },
  { key: "en_cours", label: "En cours", bg: "bg-amber-100", text: "text-amber-700" },
  { key: "termine", label: "Terminé", bg: "bg-emerald-100", text: "text-emerald-700" },
];

function ProjectsSection({ profile, projects, saveProjects, allNames, personFilter, oswald }) {
  const [modal, setModal] = useState(null);
  const [showArchive, setShowArchive] = useState(false);
  const todayISO = isoDate(new Date());

  const filtered = personFilter === "all" ? projects : projects.filter((p) => (p.assignees || []).includes(personFilter));
  const order = { a_faire: 0, en_cours: 1 };
  const active = filtered
    .filter((p) => p.status !== "termine")
    .sort((a, b) => {
      const oa = order[a.status] ?? 0;
      const ob = order[b.status] ?? 0;
      if (oa !== ob) return oa - ob;
      const overdueA = a.dueDate && a.dueDate < todayISO;
      const overdueB = b.dueDate && b.dueDate < todayISO;
      if (overdueA !== overdueB) return overdueA ? -1 : 1;
      if (a.dueDate && b.dueDate) return a.dueDate.localeCompare(b.dueDate);
      if (a.dueDate) return -1;
      if (b.dueDate) return 1;
      return 0;
    });
  const archived = filtered.filter((p) => p.status === "termine");

  async function handleSave(data) {
    let next;
    if (modal.mode === "add") {
      const id = "proj-" + Date.now();
      next = [...projects, { id, title: data.title, status: data.status, dueDate: data.dueDate, assignees: data.assignees, comment: data.comment, createdBy: profile.name, createdTs: Date.now() }];
    } else {
      next = projects.map((p) => (p.id === modal.id ? { ...p, title: data.title, status: data.status, dueDate: data.dueDate, assignees: data.assignees, comment: data.comment } : p));
    }
    await saveProjects(next);
    setModal(null);
  }
  async function handleDelete() {
    await saveProjects(projects.filter((p) => p.id !== modal.id));
    setModal(null);
  }

  return (
    <div>
      <div className="flex items-center justify-between mb-3">
        <div className="font-bold text-emerald-900" style={oswald}>Projets</div>
        <button
          onClick={() => setModal({ mode: "add", id: null, title: "", status: "a_faire", dueDate: "", assignees: [], comment: "" })}
          className="text-xs font-semibold text-white bg-emerald-600 rounded-lg px-3 py-1.5 flex items-center gap-1"
        >
          <Plus size={14} /> Nouveau projet
        </button>
      </div>

      {active.length === 0 ? (
        <div className="text-center text-slate-400 py-6 text-sm">Aucun projet en cours.</div>
      ) : (
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 divide-y divide-slate-100 mb-3">
          {active.map((p) => (
            <ProjectRow key={p.id} project={p} onEdit={() => setModal({ mode: "edit", ...p })} oswald={oswald} />
          ))}
        </div>
      )}

      {archived.length > 0 && (
        <div>
          <button onClick={() => setShowArchive((s) => !s)} className="text-xs font-semibold text-slate-400 mb-2">
            {showArchive ? "Masquer" : "Voir"} les archives ({archived.length})
          </button>
          {showArchive && (
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 divide-y divide-slate-100">
              {archived.map((p) => (
                <ProjectRow key={p.id} project={p} onEdit={() => setModal({ mode: "edit", ...p })} oswald={oswald} />
              ))}
            </div>
          )}
        </div>
      )}

      {modal && <ProjectModal modal={modal} onClose={() => setModal(null)} onSave={handleSave} onDelete={handleDelete} oswald={oswald} allNames={allNames} />}
    </div>
  );
}

function ProjectRow({ project, onEdit, oswald }) {
  const status = PROJECT_STATUSES.find((s) => s.key === project.status) || PROJECT_STATUSES[0];
  const overdue = project.dueDate && project.status !== "termine" && project.dueDate < isoDate(new Date());
  return (
    <div className={"flex items-start gap-2 px-3 py-2.5 " + (overdue ? "bg-rose-50" : "")}>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2 flex-wrap">
          <div className={"text-sm font-semibold " + (overdue ? "text-rose-700" : "text-slate-700")}>{project.title}</div>
          <span className={"text-xs font-bold rounded px-1.5 py-0.5 " + status.bg + " " + status.text}>{status.label}</span>
          {overdue && <span className="text-xs bg-rose-600 text-white font-bold rounded px-1.5 py-0.5">En retard</span>}
        </div>
        {project.dueDate && (
          <div className={"text-xs mt-1 " + (overdue ? "text-rose-600 font-semibold" : "text-slate-400")}>
            Résolution : {new Date(project.dueDate + "T00:00:00").toLocaleDateString("fr-FR", { day: "numeric", month: "short", year: "numeric" })}
          </div>
        )}
        {project.assignees && project.assignees.length > 0 && (
          <div className="flex flex-wrap gap-1.5 mt-1">
            {project.assignees.map((n) => (
              <span key={n} className="text-xs bg-slate-100 text-slate-500 rounded px-1.5 py-0.5">{n}</span>
            ))}
          </div>
        )}
        {project.comment && <div className="text-xs text-slate-500 mt-1">{project.comment}</div>}
      </div>
      <button onClick={onEdit} className="text-slate-400 shrink-0"><Pencil size={14} /></button>
    </div>
  );
}

function ProjectModal({ modal, onClose, onSave, onDelete, oswald, allNames }) {
  const [title, setTitle] = useState(modal.title);
  const [status, setStatus] = useState(modal.status);
  const [dueDate, setDueDate] = useState(modal.dueDate || "");
  const [assignees, setAssignees] = useState(modal.assignees || []);
  const [comment, setComment] = useState(modal.comment || "");
  const [error, setError] = useState("");

  function toggleAssignee(n) {
    setAssignees((prev) => (prev.includes(n) ? prev.filter((x) => x !== n) : [...prev, n]));
  }

  function submit() {
    if (!title.trim()) { setError("Le titre du projet est requis."); return; }
    onSave({ title: title.trim(), status, dueDate, assignees, comment });
  }

  return (
    <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(15, 23, 42, 0.5)", zIndex: 50 }} onClick={onClose}>
      <div className="bg-white rounded-2xl w-full p-5 overflow-y-auto" style={{ maxWidth: "420px", maxHeight: "88vh" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-4">
          <div className="font-bold text-emerald-900" style={oswald}>{modal.mode === "add" ? "Nouveau projet" : "Modifier le projet"}</div>
          <button onClick={onClose} className="text-slate-400"><X size={20} /></button>
        </div>

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Titre</label>
        <input value={title} onChange={(e) => setTitle(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" placeholder="Ex: Refaire la carte des bières" />

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Date de résolution (optionnel)</label>
        <input type="date" value={dueDate} onChange={(e) => setDueDate(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" />

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Statut</label>
        <div className="grid grid-cols-3 gap-1.5 mb-3">
          {PROJECT_STATUSES.map((s) => (
            <button key={s.key} onClick={() => setStatus(s.key)} className={"py-2 rounded-lg text-xs font-bold border " + (status === s.key ? s.bg + " " + s.text + " border-transparent" : "border-slate-200 text-slate-400")}>
              {s.label}
            </button>
          ))}
        </div>

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Assigné à (plusieurs possibles)</label>
        <div className="flex flex-wrap gap-1.5 mb-3">
          {allNames.map((n) => (
            <button key={n} onClick={() => toggleAssignee(n)} className={"text-xs font-semibold rounded-full px-2.5 py-1 border " + (assignees.includes(n) ? "bg-emerald-600 text-white border-emerald-600" : "border-slate-200 text-slate-500")}>
              {n}
            </button>
          ))}
        </div>

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Commentaire</label>
        <textarea value={comment} onChange={(e) => setComment(e.target.value)} rows={2} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" />

        {error && <div className="text-rose-600 text-xs mb-3">{error}</div>}

        <div className="flex gap-2">
          {modal.mode === "edit" && (
            <button onClick={onDelete} className="px-3 py-2 rounded-lg border border-rose-200 text-rose-600 flex items-center gap-1 text-sm font-semibold">
              <Trash2 size={16} /> Supprimer
            </button>
          )}
          <button onClick={submit} className="flex-1 py-2 rounded-lg bg-emerald-600 text-white font-bold text-sm">Enregistrer</button>
        </div>
      </div>
    </div>
  );
}

/* ---------------- STOCK ---------------- */

function StockTab({ products, saveProducts, checked, saveChecked, oswald }) {
  const [manageOpen, setManageOpen] = useState(false);
  const [modal, setModal] = useState(null);

  async function toggleChecked(id) {
    const next = checked.includes(id) ? checked.filter((x) => x !== id) : [...checked, id];
    await saveChecked(next);
  }
  async function uncheckAll() {
    await saveChecked([]);
  }

  async function saveProduct(data) {
    let next;
    if (modal.mode === "add") {
      const id = "prod-" + Date.now();
      next = [...products, { id, group: data.group, name: data.name }];
    } else {
      next = products.map((p) => (p.id === modal.id ? { ...p, group: data.group, name: data.name } : p));
    }
    await saveProducts(next);
    setModal(null);
  }
  async function deleteProduct() {
    await saveProducts(products.filter((p) => p.id !== modal.id));
    if (checked.includes(modal.id)) await saveChecked(checked.filter((x) => x !== modal.id));
    setModal(null);
  }

  const groups = useMemo(() => {
    const order = [];
    const map = {};
    products.forEach((p) => {
      if (!map[p.group]) { map[p.group] = []; order.push(p.group); }
      map[p.group].push(p);
    });
    return order.map((g) => ({ name: g, items: map[g] }));
  }, [products]);

  return (
    <div>
      <div className="flex items-center justify-between mb-3">
        <div>
          <div className="font-bold text-emerald-900" style={oswald}>Stock à commander</div>
          <div className="text-xs text-slate-400">{checked.length} produit{checked.length > 1 ? "s" : ""} manquant{checked.length > 1 ? "s" : ""}</div>
        </div>
        <button onClick={() => setManageOpen((o) => !o)} className="text-xs font-semibold text-emerald-700 border border-emerald-200 rounded-lg px-3 py-1.5">
          {manageOpen ? "Fermer la gestion" : "Gérer les produits"}
        </button>
      </div>

      <button
        onClick={uncheckAll}
        disabled={checked.length === 0}
        className="mb-3 text-xs font-semibold text-slate-500 border border-slate-200 rounded-lg px-3 py-1.5 disabled:opacity-40"
      >
        Tout décocher
      </button>

      {manageOpen && (
        <button
          onClick={() => setModal({ mode: "add", id: null, group: groups[0]?.name || "", name: "" })}
          className="w-full mb-3 py-2 rounded-lg border border-dashed border-emerald-300 text-emerald-700 text-sm font-semibold flex items-center justify-center gap-1"
        >
          <Plus size={16} /> Ajouter un produit
        </button>
      )}

      {groups.map((g) => (
        <div key={g.name} className="bg-white rounded-xl shadow-sm border border-slate-200 mb-3 overflow-hidden">
          <div className="px-3 py-2 bg-slate-50 border-b border-slate-100 font-bold text-slate-700 text-sm">{g.name}</div>
          <div className="divide-y divide-slate-100">
            {g.items.map((p) => {
              const isChecked = checked.includes(p.id);
              return (
                <div key={p.id} className="flex items-center gap-2 px-3 py-2">
                  <button
                    onClick={() => toggleChecked(p.id)}
                    className={"w-5 h-5 rounded border shrink-0 flex items-center justify-center " + (isChecked ? "bg-rose-500 border-rose-500" : "border-slate-300")}
                  >
                    {isChecked && <Check size={14} className="text-white" />}
                  </button>
                  <div className={"flex-1 text-sm " + (isChecked ? "text-rose-600 font-semibold" : "text-slate-700")}>{p.name}</div>
                  {manageOpen && (
                    <button onClick={() => setModal({ mode: "edit", id: p.id, group: p.group, name: p.name })} className="text-slate-400 shrink-0">
                      <Pencil size={14} />
                    </button>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      ))}

      {products.length === 0 && <div className="text-center text-slate-400 py-10">Aucun produit pour l'instant.</div>}

      {modal && <StockProductModal modal={modal} groups={groups.map((g) => g.name)} onClose={() => setModal(null)} onSave={saveProduct} onDelete={deleteProduct} oswald={oswald} />}
    </div>
  );
}

function StockProductModal({ modal, groups, onClose, onSave, onDelete, oswald }) {
  const [group, setGroup] = useState(modal.group);
  const [name, setName] = useState(modal.name);
  const [error, setError] = useState("");

  function submit() {
    if (!name.trim() || !group.trim()) { setError("Le groupe et le nom du produit sont requis."); return; }
    onSave({ group: group.trim(), name: name.trim() });
  }

  return (
    <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(15, 23, 42, 0.5)", zIndex: 50 }} onClick={onClose}>
      <div className="bg-white rounded-2xl w-full p-5" style={{ maxWidth: "380px" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-4">
          <div className="font-bold text-emerald-900" style={oswald}>{modal.mode === "add" ? "Nouveau produit" : "Modifier le produit"}</div>
          <button onClick={onClose} className="text-slate-400"><X size={20} /></button>
        </div>

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Groupe</label>
        <input list="stock-groups-list" value={group} onChange={(e) => setGroup(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" placeholder="Ex: Planche" />
        <datalist id="stock-groups-list">
          {groups.map((g) => <option key={g} value={g} />)}
        </datalist>

        <label className="text-xs font-semibold text-slate-500 mb-1 block">Produit</label>
        <input value={name} onChange={(e) => setName(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" placeholder="Ex: Chorizo" />

        {error && <div className="text-rose-600 text-xs mb-3">{error}</div>}

        <div className="flex gap-2">
          {modal.mode === "edit" && (
            <button onClick={onDelete} className="px-3 py-2 rounded-lg border border-rose-200 text-rose-600 flex items-center gap-1 text-sm font-semibold">
              <Trash2 size={16} /> Supprimer
            </button>
          )}
          <button onClick={submit} className="flex-1 py-2 rounded-lg bg-emerald-600 text-white font-bold text-sm">Enregistrer</button>
        </div>
      </div>
    </div>
  );
}

/* ---------------- HISTORIQUE ---------------- */

function HistoriqueTab({ oswald }) {
  const [loading, setLoading] = useState(true);
  const [days, setDays] = useState([]);
  const [typeFilter, setTypeFilter] = useState("all");
  const [nameFilter, setNameFilter] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [tempOnly, setTempOnly] = useState(false);

  useEffect(() => {
    (async () => {
      setLoading(true);
      const [tasksO, tasksF, doneO, doneF, checklistO, checklistF] = await Promise.all([
        storageGet("tasks:ouverture:v4", []),
        storageGet("tasks:fermeture:v4", []),
        storageListByPrefix("process-done:ouverture:"),
        storageListByPrefix("process-done:fermeture:"),
        storageListByPrefix("checklist:ouverture:"),
        storageListByPrefix("checklist:fermeture:"),
      ]);

      const map = {};
      function ensure(date) {
        if (!map[date]) map[date] = { date, ouverture: null, fermeture: null };
        return map[date];
      }
      function initType(rec, type) {
        if (!rec[type]) rec[type] = { finished: null, doneTasks: [], notDoneTasks: [] };
        return rec[type];
      }

      doneO.forEach(({ key, value }) => {
        const date = key.split(":")[2];
        if (!date) return;
        initType(ensure(date), "ouverture").finished = value;
      });
      doneF.forEach(({ key, value }) => {
        const date = key.split(":")[2];
        if (!date) return;
        initType(ensure(date), "fermeture").finished = value;
      });

      function applyChecklist(entries, type, tasksList) {
        const taskById = {};
        tasksList.forEach((t) => { taskById[t.id] = t; });
        entries.forEach(({ key, value }) => {
          const date = key.split(":")[2];
          if (!date) return;
          const slot = initType(ensure(date), type);
          const doneIds = new Set(Object.keys(value || {}));
          Object.entries(value || {}).forEach(([taskId, d]) => {
            if (!d) return;
            const task = taskById[taskId];
            slot.doneTasks.push({ taskId, text: task ? task.text : taskId, by: d.by, ts: d.ts, value: d.value });
          });
          tasksList.forEach((t) => {
            if (!doneIds.has(t.id)) slot.notDoneTasks.push({ taskId: t.id, text: t.text });
          });
        });
      }
      applyChecklist(checklistO, "ouverture", tasksO);
      applyChecklist(checklistF, "fermeture", tasksF);

      const list = Object.values(map).sort((a, b) => b.date.localeCompare(a.date));
      setDays(list);
      setLoading(false);
    })();
  }, []);

  const filtered = days.filter((d) => {
    if (typeFilter === "ouverture" && !d.ouverture) return false;
    if (typeFilter === "fermeture" && !d.fermeture) return false;
    if (dateFrom && d.date < dateFrom) return false;
    if (dateTo && d.date > dateTo) return false;
    if (nameFilter.trim()) {
      const q = nameFilter.trim().toLowerCase();
      const names = [
        d.ouverture && d.ouverture.finished ? d.ouverture.finished.by : null,
        d.fermeture && d.fermeture.finished ? d.fermeture.finished.by : null,
        ...((d.ouverture && d.ouverture.doneTasks) || []).map((v) => v.by),
        ...((d.fermeture && d.fermeture.doneTasks) || []).map((v) => v.by),
      ].filter(Boolean);
      if (!names.some((n) => n.toLowerCase().includes(q))) return false;
    }
    return true;
  });

  const tempEntries = [];
  days.forEach((d) => {
    ["ouverture", "fermeture"].forEach((type) => {
      const rec = d[type];
      if (!rec) return;
      rec.doneTasks.forEach((t) => {
        if (t.value !== undefined) tempEntries.push({ date: d.date, type, ...t });
      });
    });
  });
  tempEntries.sort((a, b) => b.date.localeCompare(a.date) || b.ts - a.ts);
  const filteredTemps = tempEntries.filter((t) => {
    if (dateFrom && t.date < dateFrom) return false;
    if (dateTo && t.date > dateTo) return false;
    if (nameFilter.trim() && !(t.by || "").toLowerCase().includes(nameFilter.trim().toLowerCase())) return false;
    return true;
  });

  if (loading) return <div className="text-center text-slate-400 py-10">Chargement…</div>;

  return (
    <div>
      <div className="flex items-center justify-between mb-3 gap-2">
        <div className="font-bold text-emerald-900" style={oswald}>Historique</div>
        <button
          onClick={() => setTempOnly((v) => !v)}
          className={"text-xs font-semibold rounded-lg px-3 py-1.5 border shrink-0 " + (tempOnly ? "bg-sky-600 text-white border-sky-600" : "border-sky-200 text-sky-700")}
        >
          {tempOnly ? "Tout l'historique" : "Températures uniquement"}
        </button>
      </div>

      <div className="flex flex-wrap items-center gap-2 mb-4">
        {!tempOnly && (
          <select value={typeFilter} onChange={(e) => setTypeFilter(e.target.value)} className="border border-slate-200 rounded-lg px-2 py-1.5 text-sm">
            <option value="all">Tous</option>
            <option value="ouverture">Ouverture</option>
            <option value="fermeture">Fermeture</option>
          </select>
        )}
        <div className="flex items-center gap-1">
          <input type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} className="border border-slate-200 rounded-lg px-2 py-1.5 text-sm" />
          <span className="text-xs text-slate-400">à</span>
          <input type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)} className="border border-slate-200 rounded-lg px-2 py-1.5 text-sm" />
        </div>
        <input
          value={nameFilter}
          onChange={(e) => setNameFilter(e.target.value)}
          placeholder="Filtrer par personne..."
          className="flex-1 border border-slate-200 rounded-lg px-3 py-1.5 text-sm"
        />
      </div>

      {tempOnly ? (
        filteredTemps.length === 0 ? (
          <div className="text-center text-slate-400 py-10">Aucune température enregistrée.</div>
        ) : (
          <div className="bg-white rounded-xl border border-slate-200 divide-y divide-slate-100">
            {filteredTemps.map((t, i) => (
              <div key={i} className="flex items-center justify-between px-3 py-2.5 gap-2">
                <div className="min-w-0">
                  <div className="text-sm font-semibold text-slate-700">{t.text}</div>
                  <div className="text-xs text-slate-400 capitalize">
                    {new Date(t.date + "T00:00:00").toLocaleDateString("fr-FR", { weekday: "short", day: "numeric", month: "short" })} · {t.type} · {t.by} à {new Date(t.ts).toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })}
                  </div>
                </div>
                <span className="text-sm font-bold text-sky-700 bg-sky-50 rounded px-2 py-1 shrink-0">{t.value}</span>
              </div>
            ))}
          </div>
        )
      ) : filtered.length === 0 ? (
        <div className="text-center text-slate-400 py-10">Aucun historique pour cette période.</div>
      ) : (
        <div className="space-y-3">
          {filtered.map((d) => (
            <div key={d.date} className="bg-white rounded-xl border border-slate-200 overflow-hidden">
              <div className="px-3 py-2 bg-slate-50 border-b border-slate-100 font-bold text-slate-700 text-sm capitalize">
                {new Date(d.date + "T00:00:00").toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long", year: "numeric" })}
              </div>
              <div className="divide-y divide-slate-100">
                {(typeFilter === "all" || typeFilter === "ouverture") && d.ouverture && (
                  <HistoriqueRow label="Ouverture" record={d.ouverture} />
                )}
                {(typeFilter === "all" || typeFilter === "fermeture") && d.fermeture && (
                  <HistoriqueRow label="Fermeture" record={d.fermeture} />
                )}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function HistoriqueRow({ label, record }) {
  const [open, setOpen] = useState(false);
  const total = record.doneTasks.length + record.notDoneTasks.length;
  const temps = record.doneTasks.filter((t) => t.value !== undefined);

  return (
    <div className="px-3 py-2.5">
      <div className="flex items-center justify-between gap-2">
        <span className="text-sm font-semibold text-slate-700">{label}</span>
        {record.finished ? (
          <span className="text-xs text-emerald-600 font-semibold text-right">
            Terminé par {record.finished.by} à {new Date(record.finished.ts).toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })}
          </span>
        ) : (
          <span className="text-xs text-slate-400">Non terminé</span>
        )}
      </div>

      {temps.length > 0 && (
        <div className="mt-1.5 flex flex-wrap gap-1.5">
          {temps.map((v, i) => (
            <span key={i} className="text-xs bg-sky-50 text-sky-700 rounded px-1.5 py-0.5">
              {v.text} : {v.value}
            </span>
          ))}
        </div>
      )}

      {total > 0 && (
        <button onClick={() => setOpen((o) => !o)} className="text-xs text-emerald-700 font-semibold mt-1.5">
          {open ? "Masquer" : "Voir"} le détail des tâches ({record.doneTasks.length}/{total})
        </button>
      )}

      {open && (
        <div className="mt-2 space-y-1">
          {record.doneTasks.map((t, i) => (
            <div key={"d" + i} className="flex items-center gap-1.5 text-xs">
              <Check size={12} className="text-emerald-600 shrink-0" />
              <span className="text-slate-600">{t.text}</span>
              <span className="text-slate-400">— {t.by} à {new Date(t.ts).toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })}</span>
            </div>
          ))}
          {record.notDoneTasks.map((t, i) => (
            <div key={"n" + i} className="flex items-center gap-1.5 text-xs">
              <X size={12} className="text-rose-400 shrink-0" />
              <span className="text-slate-400">{t.text}</span>
              <span className="text-rose-400">— non fait</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ---------------- CP ---------------- */





function CPTab({ markers, employees, oswald }) {
  const cpEntries = markers.filter((m) => m.type === "cp").sort((a, b) => a.date.localeCompare(b.date));
  const totals = {};
  cpEntries.forEach((m) => {
    totals[m.employee] = (totals[m.employee] || 0) + 1;
  });

  return (
    <div>
      <div className="font-bold text-emerald-900 mb-1" style={oswald}>Congés payés</div>
      <p className="text-xs text-slate-400 mb-3">
        Pour poser ou retirer un CP, utilisez le bouton "CP" à côté du nom de l'employé dans le planning.
      </p>

      {Object.keys(totals).length > 0 && (
        <div className="flex flex-wrap gap-2 mb-4">
          {employees.map((e) =>
            totals[e.name] ? (
              <div key={e.id} className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-slate-50">
                <span className={"w-2.5 h-2.5 rounded-full " + empColor(e.name, employees).bg} />
                <span className="text-sm font-semibold text-slate-700">{e.name}</span>
                <span className="text-sm font-bold text-slate-900" style={oswald}>{totals[e.name]} j</span>
              </div>
            ) : null
          )}
        </div>
      )}

      {cpEntries.length === 0 ? (
        <div className="text-center text-slate-400 py-10">Aucun CP posé pour l'instant.</div>
      ) : (
        <div className="bg-white rounded-xl border border-slate-200 divide-y divide-slate-100">
          {cpEntries.map((m) => (
            <div key={m.id} className="flex items-center justify-between px-3 py-2.5">
              <span className="text-sm font-semibold text-slate-700">{m.employee}</span>
              <span className="text-sm text-slate-500 capitalize">
                {new Date(m.date + "T00:00:00").toLocaleDateString("fr-FR", { weekday: "short", day: "numeric", month: "short", year: "numeric" })}
              </span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ---------------- RÉGLAGES ---------------- */

function SettingsTab({ employees, saveEmployees, eventGroups, saveEventGroups, saveActions, saveEventEntries, oswald }) {
  const [name, setName] = useState("");
  const [confirmId, setConfirmId] = useState(null);
  const [error, setError] = useState("");

  async function addEmployee() {
    const trimmed = name.trim();
    if (!trimmed) return;
    if (employees.some((e) => e.name.toLowerCase() === trimmed.toLowerCase())) {
      setError("Cet employé existe déjà.");
      return;
    }
    const colorIdx = employees.length % EMP_PALETTE.length;
    const next = [...employees, { id: "emp-" + Date.now(), name: trimmed, colorIdx }];
    await saveEmployees(next);
    setName("");
    setError("");
  }
  async function removeEmployee(id) {
    await saveEmployees(employees.filter((e) => e.id !== id));
    setConfirmId(null);
  }

  return (
    <div>
      <div className="font-bold text-emerald-900 mb-1" style={oswald}>Employés</div>
      <p className="text-xs text-slate-400 mb-3">Les employés ajoutés ici apparaissent dans le sélecteur de profil (pour badger) et comme lignes dans le planning.</p>

      <div className="bg-white rounded-xl border border-slate-200 divide-y divide-slate-100 mb-4">
        {employees.map((emp) => {
          const c = EMP_PALETTE[emp.colorIdx % EMP_PALETTE.length];
          return (
            <div key={emp.id} className="flex items-center justify-between px-3 py-2.5">
              <div className="flex items-center gap-2">
                <span className={"w-3 h-3 rounded-full shrink-0 " + c.bg} />
                <span className="text-sm font-semibold text-slate-700">{emp.name}</span>
              </div>
              {confirmId === emp.id ? (
                <div className="flex items-center gap-2">
                  <span className="text-xs text-slate-400">Supprimer ?</span>
                  <button onClick={() => removeEmployee(emp.id)} className="text-xs font-bold text-rose-600">Oui</button>
                  <button onClick={() => setConfirmId(null)} className="text-xs text-slate-400">Annuler</button>
                </div>
              ) : (
                <button onClick={() => setConfirmId(emp.id)} className="text-slate-400"><Trash2 size={16} /></button>
              )}
            </div>
          );
        })}
        {employees.length === 0 && <div className="px-3 py-4 text-sm text-slate-400 text-center">Aucun employé pour l'instant.</div>}
      </div>

      <div className="flex gap-2">
        <input
          value={name}
          onChange={(e) => { setName(e.target.value); setError(""); }}
          onKeyDown={(e) => { if (e.key === "Enter") addEmployee(); }}
          placeholder="Nom de l'employé"
          className="flex-1 border border-slate-200 rounded-lg px-3 py-2 text-sm"
        />
        <button onClick={addEmployee} className="px-4 py-2 rounded-lg bg-emerald-600 text-white font-bold text-sm flex items-center gap-1">
          <Plus size={16} /> Ajouter
        </button>
      </div>
      {error && <div className="text-rose-600 text-xs mt-2">{error}</div>}

      <p className="text-xs text-slate-400 mt-4">
        Supprimer un employé le retire du sélecteur de profil et des lignes du planning ; les créneaux déjà enregistrés à son nom restent en mémoire et réapparaissent si vous le rajoutez.
      </p>

      <EventGroupsSection eventGroups={eventGroups} saveEventGroups={saveEventGroups} oswald={oswald} />

      <ImportPlanningSection employees={employees} oswald={oswald} />

      <DangerZone saveActions={saveActions} saveEventEntries={saveEventEntries} oswald={oswald} />
    </div>
  );
}

function rowsFromWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, dateNF: "yyyy-mm-dd" });
  return rows.map((r) => r.map((c) => (c === undefined || c === null ? "" : String(c).trim())));
}

function parseTime(str) {
  if (!str) return null;
  const clean = str.trim().replace(/[hH]/, ":");
  const parts = clean.split(":");
  const h = parseInt(parts[0], 10);
  const m = parts[1] ? parseInt(parts[1], 10) : 0;
  if (isNaN(h)) return null;
  return h * 60 + (isNaN(m) ? 0 : m);
}

function EventGroupsSection({ eventGroups, saveEventGroups, oswald }) {
  const [name, setName] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editValue, setEditValue] = useState("");
  const [confirmId, setConfirmId] = useState(null);

  async function addGroup() {
    const trimmed = name.trim();
    if (!trimmed) return;
    if (eventGroups.some((g) => g.name.toLowerCase() === trimmed.toLowerCase())) return;
    await saveEventGroups([...eventGroups, { id: "eg-" + Date.now(), name: trimmed }]);
    setName("");
  }
  async function renameGroup(id) {
    const trimmed = editValue.trim();
    if (!trimmed) { setEditingId(null); return; }
    await saveEventGroups(eventGroups.map((g) => (g.id === id ? { ...g, name: trimmed } : g)));
    setEditingId(null);
  }
  async function removeGroup(id) {
    await saveEventGroups(eventGroups.filter((g) => g.id !== id));
    setConfirmId(null);
  }

  return (
    <div className="mt-6 pt-4 border-t border-slate-200">
      <div className="font-bold text-emerald-900 mb-1" style={oswald}>Groupes d'évènements</div>
      <p className="text-xs text-slate-400 mb-3">
        Ces groupes sont proposés lors de l'ajout d'un évènement (Event, Tournoi, Actualité...) dans l'onglet Événement.
      </p>
      <div className="bg-white rounded-xl border border-slate-200 divide-y divide-slate-100 mb-3">
        {eventGroups.map((g) => (
          <div key={g.id} className="flex items-center justify-between px-3 py-2.5">
            {editingId === g.id ? (
              <input
                autoFocus
                value={editValue}
                onChange={(e) => setEditValue(e.target.value)}
                onBlur={() => renameGroup(g.id)}
                onKeyDown={(e) => { if (e.key === "Enter") renameGroup(g.id); }}
                className="flex-1 border border-slate-200 rounded px-2 py-1 text-sm mr-2"
              />
            ) : (
              <span className="text-sm font-semibold text-slate-700">{g.name}</span>
            )}
            <div className="flex items-center gap-2 shrink-0">
              {editingId !== g.id && (
                <button onClick={() => { setEditingId(g.id); setEditValue(g.name); }} className="text-slate-400">
                  <Pencil size={14} />
                </button>
              )}
              {confirmId === g.id ? (
                <>
                  <button onClick={() => removeGroup(g.id)} className="text-xs font-bold text-rose-600">Oui</button>
                  <button onClick={() => setConfirmId(null)} className="text-xs text-slate-400">Annuler</button>
                </>
              ) : (
                <button onClick={() => setConfirmId(g.id)} className="text-slate-400"><Trash2 size={14} /></button>
              )}
            </div>
          </div>
        ))}
        {eventGroups.length === 0 && <div className="px-3 py-4 text-sm text-slate-400 text-center">Aucun groupe.</div>}
      </div>
      <div className="flex gap-2">
        <input
          value={name}
          onChange={(e) => setName(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter") addGroup(); }}
          placeholder="Nom du groupe"
          className="flex-1 border border-slate-200 rounded-lg px-3 py-2 text-sm"
        />
        <button onClick={addGroup} className="px-4 py-2 rounded-lg bg-emerald-600 text-white font-bold text-sm flex items-center gap-1">
          <Plus size={16} /> Ajouter
        </button>
      </div>
    </div>
  );
}

function ImportPlanningSection({ employees, oswald }) {
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState(null);
  const fileRef = useRef(null);

  async function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setBusy(true);
    setStatus(null);

    const buffer = await file.arrayBuffer();
    const rows = rowsFromWorkbook(buffer);
    const normalize = (s) => s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
    const header = (rows[0] || []).map((h) => normalize(h));
    const idx = {
      date: header.findIndex((h) => h.startsWith("date")),
      employee: header.findIndex((h) => h.startsWith("employ")),
      start: header.findIndex((h) => h.startsWith("debut")),
      end: header.findIndex((h) => h.startsWith("fin")),
      comment: header.findIndex((h) => h.startsWith("comm")),
    };

    const warnings = [];
    const byWeek = {};
    let count = 0;

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r.every((c) => !c)) continue;

      const dateStr = r[idx.date];
      const empName = r[idx.employee];
      const startStr = r[idx.start];
      const endStr = r[idx.end];
      const comment = idx.comment >= 0 ? r[idx.comment] || "" : "";

      const date = new Date(dateStr);
      if (!dateStr || isNaN(date.getTime())) {
        warnings.push(`Ligne ${i + 1} : date invalide ("${dateStr}")`);
        continue;
      }

      const emp = employees.find((e) => e.name.toLowerCase() === (empName || "").trim().toLowerCase());
      if (!emp) {
        warnings.push(`Ligne ${i + 1} : employé inconnu ("${empName}")`);
        continue;
      }

      const start = parseTime(startStr);
      const end = parseTime(endStr);
      if (start == null || end == null || end <= start) {
        warnings.push(`Ligne ${i + 1} : heures invalides ("${startStr}" - "${endStr}")`);
        continue;
      }

      const monday = mondayOf(date);
      const weekKey = isoDate(monday);
      const dayIndex = (date.getDay() + 6) % 7;

      if (!byWeek[weekKey]) byWeek[weekKey] = [];
      byWeek[weekKey].push({
        id: "imp-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8),
        day: dayIndex,
        employee: emp.name,
        start,
        end,
        comment,
      });
      count++;
    }

    for (const weekKey of Object.keys(byWeek)) {
      const existing = (await storageGet("planning:" + weekKey, [])) || [];
      await storageSet("planning:" + weekKey, [...existing, ...byWeek[weekKey]]);
    }

    setBusy(false);
    setStatus({ count, warnings });
    if (fileRef.current) fileRef.current.value = "";
  }

  return (
    <div className="mt-6 pt-4 border-t border-slate-200">
      <div className="font-bold text-emerald-900 mb-1" style={oswald}>Importer un planning</div>
      <p className="text-xs text-slate-400 mb-3">
        Fichier Excel (.xlsx) avec les colonnes : Date (AAAA-MM-JJ), Employé, Début (HH:MM), Fin (HH:MM), Commentaire (optionnel), sur la première feuille. Les créneaux importés s'ajoutent à ceux déjà présents, sans rien effacer.
      </p>
      <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} disabled={busy} className="text-sm" />
      {busy && <div className="text-xs text-slate-400 mt-2">Import en cours…</div>}
      {status && (
        <div className="mt-3 text-sm">
          <div className="text-emerald-700 font-semibold">
            {status.count} créneau{status.count > 1 ? "x" : ""} importé{status.count > 1 ? "s" : ""}.
          </div>
          {status.warnings.length > 0 && (
            <div className="mt-2 bg-amber-50 border border-amber-100 rounded-lg p-2 text-xs text-amber-700 space-y-0.5">
              {status.warnings.map((w, i) => (
                <div key={i}>{w}</div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
/* ---------------- ZONE DANGEREUSE ---------------- */

function DangerZone({ saveActions, saveEventEntries, oswald }) {
  async function clearPlanning() {
    await storageDeleteByPrefix("planning:");
    await storageDeleteByPrefix("events:");
  }
  async function clearEvents() {
    await saveEventEntries([]);
  }
  async function clearActions() {
    await saveActions([]);
  }

  return (
    <div className="mt-6 pt-4 border-t border-slate-200">
      <div className="font-bold text-rose-700 mb-1" style={oswald}>Zone dangereuse</div>
      <p className="text-xs text-slate-400 mb-3">
        Ces actions sont irréversibles et demandent une double confirmation.
      </p>
      <DangerButton
        label="Effacer tout le planning"
        description="Tous les créneaux et notes du planning seront supprimés, pour toutes les semaines."
        onConfirm={clearPlanning}
      />
      <DangerButton
        label="Effacer tous les évènements"
        description="Tous les anniversaires, réservations et tournois de l'onglet Événement seront supprimés."
        onConfirm={clearEvents}
      />
      <DangerButton
        label="Effacer toutes les actions"
        description="Toutes les actions de l'onglet To do seront supprimées (les projets ne sont pas touchés)."
        onConfirm={clearActions}
      />
    </div>
  );
}

function DangerButton({ label, description, onConfirm }) {
  const [step, setStep] = useState(0);
  const [done, setDone] = useState(false);

  async function handleFinalConfirm() {
    await onConfirm();
    setStep(0);
    setDone(true);
    setTimeout(() => setDone(false), 3000);
  }

  if (step === 0) {
    return (
      <div className="mb-2">
        <button onClick={() => setStep(1)} className="w-full py-2 rounded-lg border border-rose-200 text-rose-600 text-sm font-semibold">
          {label}
        </button>
        {done && <div className="text-xs text-emerald-600 mt-1">Effacé.</div>}
      </div>
    );
  }
  if (step === 1) {
    return (
      <div className="mb-2 bg-rose-50 border border-rose-200 rounded-lg p-3">
        <div className="text-sm text-rose-700 font-semibold mb-2">{description}</div>
        <div className="flex gap-2">
          <button onClick={() => setStep(0)} className="flex-1 py-1.5 rounded-lg border border-slate-200 text-slate-600 text-xs font-semibold">
            Annuler
          </button>
          <button onClick={() => setStep(2)} className="flex-1 py-1.5 rounded-lg bg-rose-500 text-white text-xs font-semibold">
            Continuer
          </button>
        </div>
      </div>
    );
  }
  return (
    <div className="mb-2 bg-rose-50 border border-rose-300 rounded-lg p-3">
      <div className="text-sm text-rose-800 font-bold mb-2">Vraiment sûr ? Toutes les données seront définitivement perdues.</div>
      <div className="flex gap-2">
        <button onClick={() => setStep(0)} className="flex-1 py-1.5 rounded-lg border border-slate-200 text-slate-600 text-xs font-semibold">
          Annuler
        </button>
        <button onClick={handleFinalConfirm} className="flex-1 py-1.5 rounded-lg bg-rose-600 text-white text-xs font-bold">
          Oui, tout effacer
        </button>
      </div>
    </div>
  );
}
