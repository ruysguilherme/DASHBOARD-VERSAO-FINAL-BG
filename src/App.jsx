import { useState, useMemo, useRef, useEffect, useCallback, createContext, useContext } from "react";
import {
  BarChart,
  Bar,
  AreaChart,
  Area,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  Legend
} from "recharts";
import * as XLSX from "xlsx";

// ─── DESIGN TOKENS ───
const C = {
  bg: '#030D1E', panel: '#071422', card: '#0B1C30', cardHover: '#0F2238',
  border: '#162B40', borderLight: '#1E3A52',
  gold: '#C9A851', goldDim: 'rgba(201,168,81,0.15)', goldMid: 'rgba(201,168,81,0.35)',
  blue: '#3E8FD4', blueDim: 'rgba(62,143,212,0.15)',
  amber: '#D4925A', amberDim: 'rgba(212,146,90,0.15)',
  green: '#35B67E', greenDim: 'rgba(53,182,126,0.15)',
  red: '#D95050', redDim: 'rgba(217,80,80,0.15)',
  purple: '#9B6FE0', cyan: '#5CB8E0',
  text: '#E4EDF8', muted: '#5A7A9A', dim: '#1E3352',
};

const ML = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
const ML_FULL = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
const TIPOS = ['Caixa', 'Cartão de Crédito'];
const PESSOAS = ['Guilherme', 'Bruna'];
const CATEGORIAS = ['Despesa Fixa','Reforma Casa','Mercado','Delivery','Lazer','Assinaturas','Diversos','Saúde','Transporte','Educação','Outros'];
const CAT_C = {
  'Despesa Fixa':'#D95050','Reforma Casa':'#3E8FD4','Mercado':'#9B6FE0',
  'Delivery':'#D4925A','Lazer':'#5CB8E0','Assinaturas':'#35B67E',
  'Diversos':'#C9A851','Saúde':'#D85A8C','Transporte':'#7DA3C9',
  'Educação':'#A485D4','Outros':'#5A7A9A',
};
const TYPE_C = { 'Caixa':C.blue, 'Cartão de Crédito':C.gold };

const YEAR = 2026;
const HIDDEN_TOKEN = 'R$ ••••••';

// Default liquidated months for 2026: Feb (1), Mar (2), Apr (3)
const DEFAULT_PAID_MONTHS = { 1: true, 2: true, 3: true };

// ─── UTILITY FUNCTIONS ───
const fmt = v => `R$ ${Number(v).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const fmtK = v => v >= 1000 ? `R$ ${(v/1000).toFixed(1)}k` : fmt(v);
const pct = (v,t) => t ? ((v/t)*100).toFixed(1)+'%' : '0%';
const uid = () => Math.random().toString(36).slice(2,11);

const formatCurrency = fmt;

const capitalize = s => s ? s.charAt(0).toUpperCase() + s.slice(1) : s;

function getCurrentMonthLabel() {
  try {
    const d = new Date();
    const txt = d.toLocaleString('pt-BR', { month: 'long', year: 'numeric' });
    return capitalize(txt);
  } catch (e) {
    const idx = new Date().getMonth();
    return `${ML_FULL[idx]} de ${YEAR}`;
  }
}

// ─── HIDE BALANCES CONTEXT ───
const HideCtx = createContext(false);
function useHideBalances() { return useContext(HideCtx); }
function maskFmt(v, hide) { return hide ? HIDDEN_TOKEN : fmt(v); }
function maskFmtK(v, hide) { return hide ? HIDDEN_TOKEN : fmtK(v); }

// ─── PAID MONTHS CONTEXT ───
const PaidCtx = createContext({ paidMonths: {}, togglePaid: () => {} });
function usePaidMonths() { return useContext(PaidCtx); }

function getMonthData(rows, monthIdx) {
  return rows.map(r => ({ ...r, monthVal: r.m[monthIdx] || 0 })).filter(r => r.monthVal > 0);
}

function calculateSettlement(rows, monthIdx) {
  const slice = monthIdx !== null
    ? rows.map(r => ({ ...r, total: r.m[monthIdx] || 0 }))
    : rows;
  const g2b = slice.filter(r=>r.q==='Guilherme'&&r.p==='Bruna').reduce((s,r)=>s+r.total,0);
  const b2g = slice.filter(r=>r.q==='Bruna'&&r.p==='Guilherme').reduce((s,r)=>s+r.total,0);
  const net = g2b - b2g;
  return { g2b, b2g, net, direction: net>0.01?'g2b':net<-0.01?'b2g':'zero' };
}

function filterTransactions(rows, { person, paymentMethod, category }) {
  return rows.filter(r => {
    if (person && person !== 'all' && r.q !== person) return false;
    if (paymentMethod && paymentMethod !== 'all' && r.t !== paymentMethod) return false;
    if (category && category !== 'all' && r.c !== category) return false;
    return true;
  });
}

// ─── INITIAL DATA ───
const INITIAL = [
  {t:'Caixa',q:'Bruna',p:'Bruna',c:'Despesa Fixa',d:'Condomínio',m:[0,0,718.37,726.12,718.37,718.37,718.37,718.37,718.37,718.37,718.37,718.37]},
  {t:'Caixa',q:'Guilherme',p:'Bruna',c:'Despesa Fixa',d:'Condomínio',m:[0,0,718.37,726.12,718.37,718.37,718.37,718.37,718.37,718.37,718.37,718.37]},
  {t:'Caixa',q:'Bruna',p:'Bruna',c:'Despesa Fixa',d:'IPTU',m:[0,814.96,414.2,428.99,407.48,407.48,407.48,407.48,407.48,407.48,407.48,407.48]},
  {t:'Caixa',q:'Guilherme',p:'Bruna',c:'Despesa Fixa',d:'IPTU',m:[0,0,414.2,428.99,407.48,407.48,407.48,407.48,407.48,407.48,407.48,407.48]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Sleep House',m:[0,0,317.66,317.66,317.66,317.66,317.66,317.66,317.66,317.66,317.66,317.66]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Reforma Casa',d:'Sleep House',m:[0,0,317.66,317.66,317.66,317.66,317.66,317.66,317.66,317.66,317.66,317.66]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Louças Duchapéu',m:[0,0,468.62,468.62,468.62,468.62,468.62,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Reforma Casa',d:'Louças Duchapéu',m:[0,0,468.62,468.62,468.62,468.62,468.62,0,0,0,0,0]},
  {t:'Caixa',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Pintor',m:[0,2250,0,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Bruna',p:'Bruna',c:'Despesa Fixa',d:'Energia',m:[0,0,64.25,0,176.755,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Bruna',c:'Despesa Fixa',d:'Energia',m:[0,0,64.25,0,176.755,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Cama',m:[0,0,249.92,249.92,249.92,249.92,249.92,249.92,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Reforma Casa',d:'Cama',m:[0,0,249.92,249.92,249.92,249.92,249.92,249.92,0,0,0,0]},
  {t:'Caixa',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Leroy Merlin',m:[0,1441.71,0,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Pintor',m:[0,1300,0,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Pintor',m:[0,1200,0,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Pintor',m:[0,1000,0,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Materiais Limpeza',m:[0,473.81,0,0,0,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Camicado',m:[0,0,65.13,65.13,65.13,65.13,65.13,65.13,65.13,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Reforma Casa',d:'Camicado',m:[0,0,65.13,65.13,65.13,65.13,65.13,65.13,65.13,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Zara Home',m:[0,0,45.15,45.15,45.015,45.15,45.15,45.15,45.15,45.15,45.15,45.15]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Reforma Casa',d:'Zara Home',m:[0,0,45.15,45.15,45.015,45.15,45.15,45.15,45.15,45.15,45.15,45.15]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Materiais Pintura',m:[0,360.6,0,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Bruna',p:'Guilherme',c:'Reforma Casa',d:'Leroy Merlin',m:[0,187.33,0,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Limpeza Estofados',m:[0,0,225,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Bruna',p:'Guilherme',c:'Reforma Casa',d:'Limpeza Estofados',m:[0,0,225,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Bruna',p:'Bruna',c:'Diversos',d:'Thamires Limpeza',m:[0,0,350,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Bruna',c:'Diversos',d:'Thamires Limpeza',m:[0,0,350,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Limpeza Estofados',m:[0,0,510,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Limpeza Estofados',m:[0,0,1190,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Bruna',p:'Bruna',c:'Diversos',d:'Thamires Limpeza',m:[0,0,350,0,0,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Guilherme',c:'Reforma Casa',d:'Limpeza Estofados',m:[0,0,153.17,153.17,153.17,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Limpeza Estofados',m:[0,0,153.17,153.17,153.17,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Guilherme',c:'Reforma Casa',d:'Maquina café',m:[0,0,54.21,54.21,54.21,54.21,54.21,54.21,54.21,54.21,54.21,54.21]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Maquina café',m:[0,0,54.21,54.21,54.21,54.21,54.21,54.21,54.21,54.21,54.21,54.21]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Materiais Pintura',m:[0,152.47,0,0,0,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Materiais Limpeza',m:[0,124.96,0,0,0,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Magalu (Eletros)',m:[0,0,0,506.68,684.175,506.68,506.68,506.68,506.68,506.68,506.68,506.68]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Reforma Casa',d:'Magalu (Eletros)',m:[0,0,0,506.68,684.175,506.68,506.68,506.68,506.68,506.68,506.68,506.68]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Magalu (TV21x)',m:[0,0,0,283.89,283.89,283.89,283.89,283.89,283.89,283.89,283.89,283.89]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Reforma Casa',d:'Magalu (TV21x)',m:[0,0,0,283.89,283.89,283.89,283.89,283.89,283.89,283.89,283.89,283.89]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Magalu Anuidade',m:[0,0,0,8,8,8,8,8,8,8,8,8]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Reforma Casa',d:'Magalu Anuidade',m:[0,0,0,8,8,8,8,8,8,8,8,8]},
  {t:'Caixa',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Rede',m:[0,0,350,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Bruna',p:'Bruna',c:'Mercado',d:'Mercado',m:[0,0,5314.88,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Despesa Fixa',d:'Energia',m:[0,0,251.52,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Leroy Merlin',m:[0,0,445.98,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Leroy Merlin',m:[0,0,714.15,0,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Mercado',d:'Oba',m:[0,0,0,556.96,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Mercado',d:'Mercado',m:[0,0,0,321.81,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Lazer',d:'Lazer',m:[0,0,0,1400,0,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Quadros',m:[0,0,0,285,285,285,285,285,285,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Reforma Casa',d:'Quadros',m:[0,0,0,285,285,285,285,285,285,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Magalu (TV diferença)',m:[0,0,0,22.445,22.445,22.445,22.445,22.445,22.445,22.445,22.445,22.445]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Guilherme',c:'Reforma Casa',d:'Magalu (TV diferença)',m:[0,0,0,22.445,22.445,22.445,22.445,22.445,22.445,22.445,22.445,22.445]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Cadeiras Reforma',m:[0,0,0,523,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Mercado',m:[0,0,0,1267.05,0,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Reforma Casa',d:'Mercado',m:[0,0,0,300,0,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Plantas',m:[0,0,0,206.14,103.06,103.06,103.06,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Mercado',d:'Mercado',m:[0,0,0,736.53,0,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Mercado',d:'Mercado',m:[0,0,0,275,0,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Reforma Casa',d:'Plantas',m:[0,0,0,0,103.06,103.06,103.06,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Sofá',m:[0,0,0,0,410.06,410.06,410.06,410.06,410.06,410.06,410.06,410.06]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Reforma Casa',d:'Sofá',m:[0,0,0,0,410.06,410.06,410.06,410.06,410.06,410.06,410.06,410.06]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Microondas',m:[0,0,0,0,105.615,105.615,105.615,105.615,105.615,105.615,105.615,105.615]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Reforma Casa',d:'Microondas',m:[0,0,0,0,105.615,105.615,105.615,105.615,105.615,105.615,105.615,105.615]},
  {t:'Caixa',q:'Bruna',p:'Guilherme',c:'Despesa Fixa',d:'Internet',m:[0,0,0,0,62.94,62.94,62.94,62.94,62.94,62.94,62.94,62.94]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Despesa Fixa',d:'Internet',m:[0,0,0,0,62.94,62.94,62.94,62.94,62.94,62.94,62.94,62.94]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Lazer',d:'Lazer',m:[0,0,0,0,450,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Mercado',d:'Mercado',m:[0,0,0,0,1761.781,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Guilherme',c:'Delivery',d:'Ifood',m:[0,0,0,0,411.685,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Bruna',p:'Guilherme',c:'Lazer',d:'Lazer',m:[0,0,0,0,450,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Bruna',p:'Guilherme',c:'Mercado',d:'Mercado',m:[0,0,0,0,1761.781,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Bruna',p:'Guilherme',c:'Delivery',d:'Ifood',m:[0,0,0,0,411.685,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Assinaturas',d:'Amazon (HBO Max)',m:[0,0,0,0,22.45,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Assinaturas',d:'Amazon (HBO Max)',m:[0,0,0,0,22.45,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Assinaturas',d:'Amazon (Telecine)',m:[0,0,0,0,14.95,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Assinaturas',d:'Amazon (Telecine)',m:[0,0,0,0,14.95,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Mercado',d:'Mercado',m:[0,0,0,0,586.26,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Mercado',d:'Mercado',m:[0,0,0,0,586.26,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Reforma Casa',d:'Leroy Merlin',m:[0,0,0,0,59.235,59.235,59.235,59.235,59.235,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Reforma Casa',d:'Leroy Merlin',m:[0,0,0,0,59.235,59.235,59.235,59.235,59.235,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Diversos',d:'Nespresso Presente Vitor',m:[0,0,0,0,26.965,26.965,26.965,26.965,26.965,26.965,26.965,26.965]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Diversos',d:'Nespresso Presente Vitor',m:[0,0,0,0,26.965,26.965,26.965,26.965,26.965,26.965,26.965,26.965]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Assinaturas',d:'Disney Plus',m:[0,0,0,0,33.45,33.45,33.45,33.45,33.45,33.45,33.45,33.45]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Assinaturas',d:'Disney Plus',m:[0,0,0,0,33.45,33.45,33.45,33.45,33.45,33.45,33.45,33.45]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Diversos',d:'Estacionamento',m:[0,0,0,0,20,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Diversos',d:'Estacionamento',m:[0,0,0,0,20,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Assinaturas',d:'Amazon (AdFree)',m:[0,0,0,0,5,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Assinaturas',d:'Amazon (AdFree)',m:[0,0,0,0,5,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Mercado',d:'Mercado',m:[0,0,0,0,34.16,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Mercado',d:'Mercado',m:[0,0,0,0,34.16,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Delivery',d:'Mcdonalds',m:[0,0,0,0,48.4,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Delivery',d:'Mcdonalds',m:[0,0,0,0,48.4,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Mercado',d:'Mercado',m:[0,0,0,0,10.995,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Mercado',d:'Mercado',m:[0,0,0,0,10.995,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Mercado',d:'Feira Prédio',m:[0,0,0,0,143.5,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Mercado',d:'Feira Prédio',m:[0,0,0,0,143.5,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Assinaturas',d:'Amazon (Music Alexa)',m:[0,0,0,0,5.95,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Assinaturas',d:'Amazon (Music Alexa)',m:[0,0,0,0,5.95,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Assinaturas',d:'Netflix',m:[0,0,0,0,34.9,34.9,34.9,34.9,34.9,34.9,34.9,34.9]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Assinaturas',d:'Netflix',m:[0,0,0,0,34.9,34.9,34.9,34.9,34.9,34.9,34.9,34.9]},
  {t:'Caixa',q:'Bruna',p:'Bruna',c:'Lazer',d:'Lazer',m:[0,0,0,0,270,0,0,0,0,0,0,0]},
  {t:'Caixa',q:'Guilherme',p:'Bruna',c:'Lazer',d:'Lazer',m:[0,0,0,0,270,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Delivery',d:'Ifood',m:[0,0,0,0,1745.58,0,0,0,0,0,0,0]},
  {t:'Cartão de Crédito',q:'Guilherme',p:'Bruna',c:'Delivery',d:'Ifood',m:[0,0,0,0,1745.58,0,0,0,0,0,0,0]},
].map(r => ({ ...r, id: uid(), total: r.m.reduce((s,v)=>s+v,0) }));

// ─── STORAGE (data + UI prefs) ───
const STORAGE_KEY = 'expenses_dashboard_v3';
const HIDE_BAL_KEY = 'bg_finance_hide_balances_v1';
const PAID_MONTHS_KEY = 'bg_finance_paid_months_v1';

async function loadData() {
  try { if (window?.storage) { const r = await window.storage.get(STORAGE_KEY); if (r?.value) return JSON.parse(r.value); } } catch(e) {}
  return INITIAL;
}
async function saveData(data) {
  try { if (window?.storage) await window.storage.set(STORAGE_KEY, JSON.stringify(data)); } catch(e) {}
}

function loadHideBalancesPref() {
  try { return localStorage.getItem(HIDE_BAL_KEY) === '1'; } catch(e) { return false; }
}
function saveHideBalancesPref(v) {
  try { localStorage.setItem(HIDE_BAL_KEY, v ? '1' : '0'); } catch(e) {}
}
function loadPaidMonths() {
  try {
    const raw = localStorage.getItem(PAID_MONTHS_KEY);
    if (raw) return JSON.parse(raw);
  } catch(e) {}
  return { ...DEFAULT_PAID_MONTHS };
}
function savePaidMonths(pm) {
  try { localStorage.setItem(PAID_MONTHS_KEY, JSON.stringify(pm)); } catch(e) {}
}

// ─── ANALYTICS ENGINE ───
function recompute(rows) {
  const v = rows.map(r => ({ ...r, total: r.m.reduce((s,x)=>s+x,0) }));
  const totalGeral = v.reduce((s,r)=>s+r.total,0);
  const gTotal = v.filter(r=>r.q==='Guilherme').reduce((s,r)=>s+r.total,0);
  const bTotal = v.filter(r=>r.q==='Bruna').reduce((s,r)=>s+r.total,0);
  const g2b = v.filter(r=>r.q==='Guilherme'&&r.p==='Bruna').reduce((s,r)=>s+r.total,0);
  const b2g = v.filter(r=>r.q==='Bruna'&&r.p==='Guilherme').reduce((s,r)=>s+r.total,0);

  const monthly = ML.map((mes,i) => {
    const G = v.filter(r=>r.q==='Guilherme').reduce((s,r)=>s+(r.m[i]||0),0);
    const B = v.filter(r=>r.q==='Bruna').reduce((s,r)=>s+(r.m[i]||0),0);
    const gOwesB = v.filter(r=>r.q==='Guilherme'&&r.p==='Bruna').reduce((s,r)=>s+(r.m[i]||0),0);
    const bOwesG = v.filter(r=>r.q==='Bruna'&&r.p==='Guilherme').reduce((s,r)=>s+(r.m[i]||0),0);
    const net = gOwesB - bOwesG;

    const catMap = {};
    v.forEach(r => { catMap[r.c] = (catMap[r.c]||0)+(r.m[i]||0); });
    const topCats = Object.entries(catMap).filter(([,val])=>val>0).sort((a,b)=>b[1]-a[1]).slice(0,5);
    const typeMap = {};
    v.forEach(r => { typeMap[r.t] = (typeMap[r.t]||0)+(r.m[i]||0); });

    return {
      mes, idx:i,
      Guilherme: Math.round(G*100)/100,
      Bruna: Math.round(B*100)/100,
      total: Math.round((G+B)*100)/100,
      gOwesB: Math.round(gOwesB*100)/100,
      bOwesG: Math.round(bOwesG*100)/100,
      net: Math.round(net*100)/100,
      direction: net > 0.01 ? 'g2b' : net < -0.01 ? 'b2g' : 'zero',
      topCats, typeMap,
    };
  });

  const monthlyActive = monthly.filter(m => m.total > 0);
  const byType = {};
  v.forEach(r => { byType[r.t] = (byType[r.t]||0) + r.total; });
  const byCat = {};
  v.forEach(r => { byCat[r.c] = (byCat[r.c]||0) + r.total; });
  const byDesc = {};
  v.forEach(r => { byDesc[r.d] = (byDesc[r.d]||0) + r.total; });
  const topDesc = Object.entries(byDesc).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const fixedTotal = v.filter(r=>r.c==='Despesa Fixa').reduce((s,r)=>s+r.total,0);
  const varTotal = totalGeral - fixedTotal;

  return { rows:v, totalGeral, gTotal, bTotal, g2b, b2g, monthly, monthlyActive, byType, byCat, topDesc, fixedTotal, varTotal };
}

// Net pendente (anual) considerando meses já liquidados
function pendingSettlement(monthly, paidMonths) {
  let g2b = 0, b2g = 0;
  monthly.forEach(m => {
    if (paidMonths[m.idx]) return;
    g2b += m.gOwesB;
    b2g += m.bOwesG;
  });
  const net = g2b - b2g;
  return { g2b, b2g, net, direction: net>0.01?'g2b':net<-0.01?'b2g':'zero' };
}

// ─── SHARED STYLES ───
const s = {
  card: { background:C.card, border:`1px solid ${C.border}`, borderRadius:14, padding:'20px 24px' },
  cardSm: { background:C.card, border:`1px solid ${C.border}`, borderRadius:12, padding:'16px 20px' },
  label: { fontSize:11, letterSpacing:2, textTransform:'uppercase', color:C.muted, fontWeight:600 },
  val: { fontSize:28, fontWeight:700, fontFamily:'monospace', color:C.text },
  muted: { color:C.muted, fontSize:13 },
  tag: (col) => ({ background:`${col}22`, color:col, fontSize:11, fontWeight:600, padding:'2px 8px', borderRadius:20, letterSpacing:1, display:'inline-block' }),
  input: { background:C.panel, border:`1px solid ${C.border}`, borderRadius:8, padding:'10px 14px', color:C.text, fontSize:13, outline:'none', width:'100%', fontFamily:'inherit' },
  btn: (color=C.gold) => ({ background:color, color:C.bg, border:'none', borderRadius:8, padding:'10px 20px', fontSize:13, fontWeight:700, cursor:'pointer', transition:'all 0.2s' }),
  btnGhost: { background:'transparent', color:C.muted, border:`1px solid ${C.border}`, borderRadius:8, padding:'8px 16px', fontSize:12, cursor:'pointer', transition:'all 0.2s' },
};

// ─── TOOLTIP ───
const TT = ({ active, payload, label }) => {
  const hide = useHideBalances();
  return active && payload?.length ? (
    <div style={{ background:C.panel, border:`1px solid ${C.border}`, borderRadius:8, padding:'10px 14px', boxShadow:'0 8px 24px rgba(0,0,0,0.4)' }}>
      <p style={{ color:C.muted, fontSize:12, marginBottom:6 }}>{label}</p>
      {payload.map((p,i) => (
        <p key={i} style={{ color:p.color||C.text, fontSize:13, marginBottom:2 }}>
          {p.name}: {maskFmt(p.value, hide)}
        </p>
      ))}
    </div>
  ) : null;
};

// ─── KPI CARD ───
function KPI({ label, value, sub, accent, icon, onClick }) {
  return (
    <button onClick={onClick} style={{ ...s.card, borderTop:`3px solid ${accent}`, position:'relative', overflow:'hidden', width:'100%', textAlign:'left', cursor:onClick?'pointer':'default' }}>
      <div style={{ position:'absolute', right:16, top:16, fontSize:28, opacity:0.12 }}>{icon}</div>
      <p style={s.label}>{label}</p>
      <p style={{ ...s.val, color:accent, marginTop:6, fontSize:'clamp(18px,3vw,28px)' }}>{value}</p>
      {sub && <p style={{ ...s.muted, marginTop:4, fontSize:11 }}>{sub}</p>}
    </button>
  );
}

function Field({ label, children, span=1 }) {
  return (
    <div style={{ gridColumn:`span ${span}` }}>
      <label style={{ ...s.label, display:'block', marginBottom:6 }}>{label}</label>
      {children}
    </div>
  );
}

// ─── BADGE ───
function Badge({ text, color }) {
  return (
    <span style={{ background:`${color}22`, color, fontSize:10, fontWeight:700, padding:'3px 8px', borderRadius:20, letterSpacing:1, display:'inline-flex', alignItems:'center', gap:4, whiteSpace:'nowrap' }}>
      {text}
    </span>
  );
}

// ─── HIDE BALANCES TOGGLE (eye icon) ───
function HideBalancesToggle({ hide, onChange }) {
  return (
    <button
      onClick={() => onChange(!hide)}
      title={hide ? 'Mostrar valores' : 'Ocultar valores'}
      aria-label={hide ? 'Mostrar valores' : 'Ocultar valores'}
      aria-pressed={hide}
      style={{
        background: hide ? C.goldDim : 'transparent',
        color: hide ? C.gold : C.muted,
        border: `1px solid ${hide ? C.gold : C.border}`,
        borderRadius: 26,
        padding: '8px 14px',
        fontSize: 13,
        fontWeight: 600,
        cursor: 'pointer',
        whiteSpace: 'nowrap',
        display: 'inline-flex',
        alignItems: 'center',
        gap: 6,
        transition: 'all 0.2s',
      }}
    >
      <span aria-hidden="true">{hide ? '🙈' : '👁'}</span>
      <span>{hide ? 'Mostrar saldos' : 'Ocultar saldos'}</span>
    </button>
  );
}

// ─── VIEW TOGGLE (Mensal / Anual) ───
function ViewToggle({ viewMode, onChange }) {
  return (
    <div style={{ display:'inline-flex', background:C.panel, border:`1px solid ${C.border}`, borderRadius:30, padding:3, position:'relative' }}>
      <div style={{
        position:'absolute', top:3, left:viewMode==='monthly'?3:'calc(50% + 1px)',
        width:'calc(50% - 4px)', height:'calc(100% - 6px)',
        background:C.goldDim, border:`1px solid ${C.gold}`, borderRadius:26,
        transition:'left 0.3s cubic-bezier(0.4,0,0.2,1)',
      }}/>
      {['monthly','annual'].map(mode => (
        <button key={mode} onClick={() => onChange(mode)} style={{
          position:'relative', zIndex:1, padding:'8px 18px', borderRadius:26, border:'none',
          background:'transparent', fontSize:12, fontWeight:700, cursor:'pointer', letterSpacing:1,
          color: viewMode===mode ? C.gold : C.muted, transition:'color 0.2s',
          whiteSpace:'nowrap',
        }}>
          {mode==='monthly' ? '📅 Mensal' : '📊 Anual'}
        </button>
      ))}
    </div>
  );
}

// ─── HIDE PREVIOUS MONTHS TOGGLE ───
function HidePreviousToggle({ hidePreviousMonths, onChange, currentMonth }) {
  const disabled = currentMonth === 0;
  return (
    <button
      onClick={() => !disabled && onChange(!hidePreviousMonths)}
      disabled={disabled}
      title={disabled ? 'Não há meses anteriores em Janeiro' : ''}
      style={{
        background: hidePreviousMonths ? C.goldDim : 'transparent',
        color: disabled ? C.dim : hidePreviousMonths ? C.gold : C.muted,
        border: `1px solid ${disabled ? C.dim : hidePreviousMonths ? C.gold : C.border}`,
        borderRadius: 26,
        padding: '8px 16px',
        fontSize: 12,
        fontWeight: 700,
        cursor: disabled ? 'not-allowed' : 'pointer',
        letterSpacing: 1,
        whiteSpace: 'nowrap',
        transition: 'all 0.2s',
        opacity: disabled ? 0.5 : 1,
      }}
    >
      {hidePreviousMonths ? '👁 Mostrar meses anteriores' : '🙈 Ocultar meses anteriores'}
    </button>
  );
}

// ─── MONTH SELECTOR ───
function MonthSelector({ selectedMonth, onChange, monthly, hidePreviousMonths, currentMonth, paidMonths }) {
  const scrollRef = useRef();
  const visibleMonths = ML
    .map((mes, i) => ({ mes, i }))
    .filter(({ i }) => !hidePreviousMonths || i >= currentMonth);

  return (
    <div style={{ overflowX:'auto', WebkitOverflowScrolling:'touch', scrollbarWidth:'none', msOverflowStyle:'none' }} ref={scrollRef}>
      <div style={{ display:'flex', gap:6, padding:'4px 2px', minWidth:'max-content' }}>
        {visibleMonths.map(({ mes, i }) => {
          const mData = monthly[i];
          const hasData = mData?.total > 0;
          const isSelected = selectedMonth === i;
          const isPaid = !!paidMonths[i];
          const isCurrent = i === currentMonth;
          return (
            <button key={i} onClick={() => hasData && onChange(i)}
              aria-pressed={isSelected}
              style={{
                padding:'8px 14px', borderRadius:20,
                border:`1px solid ${isSelected ? C.gold : isPaid ? C.green : hasData ? C.border : C.dim}`,
                background: isSelected ? C.goldDim : isPaid ? C.greenDim : 'transparent',
                color: isSelected ? C.gold : isPaid ? C.green : hasData ? C.text : C.dim,
                fontSize:12, fontWeight:isSelected?700:isCurrent?700:400,
                cursor:hasData?'pointer':'not-allowed',
                transition:'all 0.2s', whiteSpace:'nowrap', position:'relative',
                minWidth:62,
              }}>
              {mes}{isCurrent ? ' •' : ''}
              {isPaid && (
                <span style={{ marginLeft:4, fontSize:10 }}>✓</span>
              )}
            </button>
          );
        })}
      </div>
    </div>
  );
}

// ─── FILTER BAR ───
function FilterBar({ filters, onChange, rows }) {
  const cats = [...new Set(rows.map(r=>r.c))].sort();
  return (
    <div style={{ display:'flex', gap:8, flexWrap:'wrap', alignItems:'center' }}>
      <span style={{ fontSize:11, color:C.muted, letterSpacing:1, textTransform:'uppercase' }}>Filtrar:</span>
      <select value={filters.person} onChange={e=>onChange({...filters,person:e.target.value})}
        style={{ ...s.input, width:'auto', fontSize:12, padding:'6px 10px', cursor:'pointer' }}>
        <option value="all">👥 Todos</option>
        {PESSOAS.map(p=><option key={p} value={p}>👤 {p}</option>)}
      </select>
      <select value={filters.paymentMethod} onChange={e=>onChange({...filters,paymentMethod:e.target.value})}
        style={{ ...s.input, width:'auto', fontSize:12, padding:'6px 10px', cursor:'pointer' }}>
        <option value="all">💳 Todos Pagamentos</option>
        {TIPOS.map(t=><option key={t} value={t}>{t}</option>)}
      </select>
      <select value={filters.category} onChange={e=>onChange({...filters,category:e.target.value})}
        style={{ ...s.input, width:'auto', fontSize:12, padding:'6px 10px', cursor:'pointer' }}>
        <option value="all">🏷 Todas Categorias</option>
        {cats.map(c=><option key={c} value={c}>{c}</option>)}
      </select>
      {(filters.person!=='all'||filters.paymentMethod!=='all'||filters.category!=='all') && (
        <button onClick={()=>onChange({person:'all',paymentMethod:'all',category:'all'})}
          style={{ ...s.btnGhost, padding:'6px 12px', fontSize:11, color:C.red, borderColor:C.red }}>
          ✕ Limpar
        </button>
      )}
    </div>
  );
}

// ─── MONTHLY SUMMARY BANNER ───
// Mostra o cabeçalho do mês selecionado com saldos, status de liquidação
// e botão para o usuário marcar o mês corrente como liquidado.
function MonthlySummaryBanner({ mData, rows, currentMonth }) {
  const hide = useHideBalances();
  const { paidMonths, togglePaid } = usePaidMonths();
  const isPaid = !!paidMonths[mData.idx];
  const isCurrent = mData.idx === currentMonth;

  const { gOwesB, bOwesG, net, direction, topCats, total } = mData;
  const dirColor = direction==='g2b'?C.red:direction==='b2g'?C.green:C.muted;
  const dirText = direction==='g2b'?'Guilherme → Bruna':direction==='b2g'?'Bruna → Guilherme':'Zerado';
  const dominantCat = topCats?.[0];
  const peakRows = rows.filter(r=>r.m[mData.idx]>0);
  const maxRow = peakRows.reduce((p,c)=>(c.m[mData.idx]||0)>(p.m[mData.idx]||0)?c:p,peakRows[0]);

  const toPayG = isPaid ? 0 : gOwesB; // Guilherme a pagar
  const toPayB = isPaid ? 0 : bOwesG; // Bruna a pagar
  const effectiveNet = isPaid ? 0 : net;

  const handleMarkPaid = () => {
    if (window.confirm(`Confirma a liquidação de ${mData.mes}/${YEAR}?\n\nO saldo devido do mês será zerado.`)) {
      togglePaid(mData.idx, true);
    }
  };
  const handleUnmark = () => {
    if (window.confirm(`Reabrir ${mData.mes}/${YEAR} como pendente?`)) {
      togglePaid(mData.idx, false);
    }
  };

  return (
    <div style={{
      ...s.card,
      background:`linear-gradient(135deg, ${C.card} 0%, #0a1f38 100%)`,
      border:`1px solid ${isPaid ? C.green : C.borderLight}`,
      marginBottom:20
    }}>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:16, flexWrap:'wrap', gap:8 }}>
        <div>
          <div style={{ display:'flex', alignItems:'center', gap:10, flexWrap:'wrap' }}>
            <p style={{ fontSize:24, fontWeight:700, color:C.text }}>{mData.mes} {YEAR}</p>
            {isCurrent && <Badge text="MÊS ATUAL" color={C.gold} />}
            {isPaid && <Badge text="✓ LIQUIDADO" color={C.green} />}
          </div>
          <p style={{ ...s.muted, fontSize:12, marginTop:4 }}>{peakRows.length} lançamentos no mês</p>
        </div>
        <div style={{ display:'flex', gap:6, flexWrap:'wrap', alignItems:'center' }}>
          {dominantCat && <Badge text={`📊 ${dominantCat[0]}`} color={CAT_C[dominantCat[0]]||C.muted} />}
          {maxRow && <Badge text={`💥 Maior: ${maxRow.d}`} color={C.gold} />}
          {!isPaid && <Badge text={`⚖️ ${dirText}`} color={dirColor} />}
        </div>
      </div>

      <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(160px,1fr))', gap:12 }}>
        {[
          { label:'Total do Mês', val:maskFmt(total, hide), color:C.text, icon:'💰' },
          { label:'Guilherme a pagar', val:maskFmt(toPayG, hide), color:C.blue, icon:'👤' },
          { label:'Bruna a pagar', val:maskFmt(toPayB, hide), color:C.amber, icon:'👤' },
          { label:'Saldo a Liquidar', val:maskFmt(Math.abs(effectiveNet), hide), color:isPaid?C.green:dirColor, icon:isPaid?'✓':'⚖️' },
        ].map(({ label,val,color,icon }) => (
          <div key={label} style={{ background:C.panel, padding:'12px 14px', borderRadius:10 }}>
            <p style={{ fontSize:11, color:C.muted, letterSpacing:1, textTransform:'uppercase', marginBottom:4 }}>{icon} {label}</p>
            <p style={{ fontSize:16, fontFamily:'monospace', fontWeight:700, color }}>{val}</p>
          </div>
        ))}
      </div>

      {isPaid ? (
        <div style={{ marginTop:14, padding:'10px 14px', background:C.greenDim, border:`1px solid ${C.green}40`, borderRadius:8, display:'flex', alignItems:'center', justifyContent:'space-between', gap:10, flexWrap:'wrap' }}>
          <div style={{ display:'flex', alignItems:'center', gap:10 }}>
            <span style={{ fontSize:18 }}>✅</span>
            <span style={{ color:C.green, fontWeight:700, fontSize:13 }}>Mês liquidado</span>
            <span style={{ color:C.muted, fontSize:12 }}>— saldo devido foi quitado.</span>
          </div>
          <button onClick={handleUnmark} style={{ ...s.btnGhost, padding:'6px 14px', fontSize:11 }} aria-label={`Reabrir ${mData.mes} como pendente`}>
            ↺ Reabrir mês
          </button>
        </div>
      ) : Math.abs(net) > 0.01 ? (
        <div style={{ marginTop:14, padding:'10px 14px', background:`${dirColor}15`, border:`1px solid ${dirColor}40`, borderRadius:8, display:'flex', alignItems:'center', justifyContent:'space-between', gap:10, flexWrap:'wrap' }}>
          <div style={{ display:'flex', alignItems:'center', gap:10 }}>
            <span style={{ fontSize:18 }}>{direction==='g2b'?'🔴':'🟢'}</span>
            <div>
              <span style={{ color:dirColor, fontWeight:700, fontSize:13 }}>{dirText}</span>
              <span style={{ color:C.muted, fontSize:12 }}> — transferência de </span>
              <span style={{ color:dirColor, fontFamily:'monospace', fontWeight:700, fontSize:13 }}>{maskFmt(Math.abs(net), hide)}</span>
            </div>
          </div>
          <button
            onClick={handleMarkPaid}
            style={s.btn(C.green)}
            aria-label={`Marcar ${mData.mes} como liquidado`}
          >
            ✓ Marcar como liquidado
          </button>
        </div>
      ) : (
        <div style={{ marginTop:14, padding:'10px 14px', background:C.panel, borderRadius:8, color:C.muted, fontSize:12 }}>
          ✅ Mês sem saldo a liquidar.
        </div>
      )}
    </div>
  );
}

// ─── TAB: VISÃO GERAL ───
function OverviewTab({ a, viewMode, selectedMonth, filters, currentMonth, onNavigate }) {
  const hide = useHideBalances();
  const { paidMonths } = usePaidMonths();
  const filteredRows = useMemo(() => filterTransactions(a.rows, filters), [a.rows, filters]);

  if (viewMode === 'monthly') {
    const mData = a.monthly[selectedMonth];
    const isPaid = !!paidMonths[selectedMonth];
    const monthRows = filteredRows.map(r => ({ ...r, val: r.m[selectedMonth]||0 })).filter(r=>r.val>0);
    const catData = {};
    monthRows.forEach(r => { catData[r.c] = (catData[r.c]||0)+r.val; });
    const catArr = Object.entries(catData).map(([name,value])=>({name,value,fill:CAT_C[name]||C.muted})).sort((a,b)=>b.value-a.value);
    const typeData = {};
    monthRows.forEach(r => { typeData[r.t] = (typeData[r.t]||0)+r.val; });
    const typeArr = Object.entries(typeData).map(([name,value])=>({name,value}));

    // Painéis "a pagar" do mês — usam o que cada pessoa precisa transferir.
    const guiAPagar = isPaid ? 0 : mData.gOwesB;
    const bruAPagar = isPaid ? 0 : mData.bOwesG;
    const compareData = [
      { name:'Guilherme', value:guiAPagar, fill:C.blue },
      { name:'Bruna',     value:bruAPagar, fill:C.amber },
    ];

    return (
      <div>
        <MonthlySummaryBanner mData={mData} rows={filteredRows} currentMonth={currentMonth} />

        <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(280px,1fr))', gap:16, marginBottom:16 }}>
          <div style={s.card}>
            <p style={{ ...s.label, marginBottom:12 }}>Top Categorias do Mês</p>
            {catArr.slice(0,5).map(({ name,value }) => (
              <div key={name} style={{ marginBottom:10 }}>
                <div style={{ display:'flex', justifyContent:'space-between', marginBottom:4 }}>
                  <span style={{ fontSize:12, color:CAT_C[name]||C.muted }}>{name}</span>
                  <span style={{ fontSize:12, fontFamily:'monospace', color:C.text }}>{maskFmt(value, hide)}</span>
                </div>
                <div style={{ height:4, background:C.dim, borderRadius:2 }}>
                  <div style={{ height:'100%', width:`${(value/(catArr[0]?.value||1))*100}%`, background:CAT_C[name]||C.muted, borderRadius:2 }}/>
                </div>
              </div>
            ))}
          </div>

          <div style={s.card}>
            <p style={{ ...s.label, marginBottom:8 }}>Forma de Pagamento</p>
            <ResponsiveContainer width="100%" height={150}>
              <PieChart>
                <Pie data={typeArr} cx="50%" cy="50%" innerRadius={40} outerRadius={65} paddingAngle={3} dataKey="value">
                  {typeArr.map((e,i) => <Cell key={i} fill={TYPE_C[e.name]||C.purple} />)}
                </Pie>
                <Tooltip content={<TT/>} />
              </PieChart>
            </ResponsiveContainer>
            <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
              {typeArr.map((e,i) => {
                const tot = typeArr.reduce((s,x)=>s+x.value,0);
                return (
                  <div key={i} style={{ display:'flex', justifyContent:'space-between' }}>
                    <span style={{ display:'flex', alignItems:'center', gap:6, fontSize:12, color:C.muted }}>
                      <span style={{ width:8,height:8,borderRadius:2,background:TYPE_C[e.name]||C.purple,display:'inline-block' }}/>
                      {e.name}
                    </span>
                    <span style={{ fontSize:12, color:C.text }}>{pct(e.value,tot)}</span>
                  </div>
                );
              })}
            </div>
          </div>

          <div style={s.card}>
            <p style={{ ...s.label, marginBottom:12 }}>A pagar — Guilherme vs Bruna</p>
            <ResponsiveContainer width="100%" height={150}>
              <BarChart data={compareData} margin={{top:0,right:10,bottom:0,left:0}}>
                <XAxis dataKey="name" tick={{fill:C.muted,fontSize:11}} axisLine={false} tickLine={false} />
                <YAxis tick={{fill:C.muted,fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>maskFmtK(v, hide)} />
                <Tooltip content={<TT/>} />
                <Bar dataKey="value" radius={[6,6,0,0]}>
                  {compareData.map((e,i)=><Cell key={i} fill={e.fill} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
            <div style={{ display:'flex', justifyContent:'space-around', marginTop:8 }}>
              <div style={{ textAlign:'center' }}>
                <p style={{ fontSize:10, color:C.muted }}>Guilherme a pagar</p>
                <p style={{ fontSize:14, fontFamily:'monospace', color:C.blue }}>{maskFmt(guiAPagar, hide)}</p>
              </div>
              <div style={{ textAlign:'center' }}>
                <p style={{ fontSize:10, color:C.muted }}>Bruna a pagar</p>
                <p style={{ fontSize:14, fontFamily:'monospace', color:C.amber }}>{maskFmt(bruAPagar, hide)}</p>
              </div>
            </div>
          </div>
        </div>

        <div style={s.card}>
          <p style={{ ...s.label, marginBottom:12 }}>Despesas do Mês</p>
          <div style={{ display:'flex', flexDirection:'column', gap:8, maxHeight:400, overflowY:'auto' }}>
            {monthRows.sort((a,b)=>b.val-a.val).map(r => (
              <div key={r.id} style={{ display:'flex', justifyContent:'space-between', alignItems:'center', padding:'8px 12px', background:C.panel, borderRadius:8, borderLeft:`3px solid ${CAT_C[r.c]||C.muted}` }}>
                <div style={{ minWidth:0 }}>
                  <p style={{ fontSize:13, color:C.text, whiteSpace:'nowrap', overflow:'hidden', textOverflow:'ellipsis' }}>{r.d}</p>
                  <div style={{ display:'flex', gap:6, marginTop:3, flexWrap:'wrap' }}>
                    <span style={s.tag(CAT_C[r.c]||C.muted)}>{r.c}</span>
                    <span style={s.tag(r.q==='Guilherme'?C.blue:C.amber)}>{r.q}</span>
                  </div>
                </div>
                <p style={{ fontSize:13, fontFamily:'monospace', color:C.text, marginLeft:12, whiteSpace:'nowrap' }}>{maskFmt(r.val, hide)}</p>
              </div>
            ))}
            {monthRows.length === 0 && <p style={{ ...s.muted, textAlign:'center', padding:20 }}>Nenhuma despesa com os filtros selecionados.</p>}
          </div>
        </div>
      </div>
    );
  }

  // Visão anual
  const pending = pendingSettlement(a.monthly, paidMonths);
  const ytg = a.monthly.slice(currentMonth + 1).reduce((s,m)=>s+m.net,0);
  const typeData = Object.entries(a.byType).map(([name,value])=>({name,value}));
  const catData = Object.entries(a.byCat).sort((x,y)=>y[1]-x[1]).map(([name,value])=>({name,value,fill:CAT_C[name]||C.muted}));

  return (
    <div>
      <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))', gap:16, marginBottom:24 }}>
        <KPI label="Total Anual" value={maskFmtK(a.totalGeral, hide)} sub={`${a.rows.length} lançamentos • clique para tabela`} accent={C.gold} icon="◈" onClick={()=>onNavigate?.('tabela')} />
        <KPI label="Guilherme" value={maskFmtK(a.gTotal, hide)} sub={pct(a.gTotal,a.totalGeral)+' do total • detalhe em encontro'} accent={C.blue} icon="◉" onClick={()=>onNavigate?.('encontro')} />
        <KPI label="Bruna" value={maskFmtK(a.bTotal, hide)} sub={pct(a.bTotal,a.totalGeral)+' do total • detalhe em encontro'} accent={C.amber} icon="◉" onClick={()=>onNavigate?.('encontro')} />
        <KPI
          label="Saldo a Liquidar"
          value={maskFmtK(Math.abs(pending.net), hide)}
          sub={pending.direction==='g2b'?'Guilherme → Bruna':pending.direction==='b2g'?'Bruna → Guilherme':'Zerado'}
          accent={pending.direction==='g2b'?C.red:pending.direction==='b2g'?C.green:C.muted}
          icon="◬"
          onClick={()=>onNavigate?.('encontro')}
        />
        <KPI label="YTG (Saldo Futuro)" value={maskFmtK(Math.abs(ytg), hide)} sub={ytg>=0?'Guilherme → Bruna':'Bruna → Guilherme'} accent={ytg>=0?C.red:C.green} icon="⌁" onClick={()=>onNavigate?.('encontro')} />
      </div>

      <div style={{ display:'grid', gridTemplateColumns:'2fr 1fr', gap:16, marginBottom:16 }}>
        <div style={s.card}>
          <p style={{ ...s.label, marginBottom:16 }}>Despesas Mensais por Pessoa</p>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={a.monthly} margin={{top:0,right:0,bottom:0,left:10}}>
              <XAxis dataKey="mes" tick={{fill:C.muted,fontSize:11}} axisLine={false} tickLine={false} />
              <YAxis tick={{fill:C.muted,fontSize:11}} axisLine={false} tickLine={false} tickFormatter={v=>hide?'•••':`${(v/1000).toFixed(0)}k`} />
              <Tooltip content={<TT/>} />
              <Bar dataKey="Guilherme" stackId="a" fill={C.blue} />
              <Bar dataKey="Bruna" stackId="a" fill={C.amber} radius={[4,4,0,0]} />
            </BarChart>
          </ResponsiveContainer>
          <div style={{ display:'flex', gap:20, marginTop:8, justifyContent:'center' }}>
            {[['Guilherme',C.blue],['Bruna',C.amber]].map(([n,c])=>(
              <span key={n} style={{ display:'flex', alignItems:'center', gap:6, fontSize:12, color:C.muted }}>
                <span style={{ width:10,height:10,borderRadius:2,background:c,display:'inline-block' }}/>{n}
              </span>
            ))}
          </div>
        </div>
        <div style={s.card}>
          <p style={{ ...s.label, marginBottom:16 }}>Forma de Pagamento</p>
          <ResponsiveContainer width="100%" height={180}>
            <PieChart>
              <Pie data={typeData} cx="50%" cy="50%" innerRadius={50} outerRadius={80} paddingAngle={3} dataKey="value">
                {typeData.map((e,i) => <Cell key={i} fill={TYPE_C[e.name]||C.purple} />)}
              </Pie>
              <Tooltip content={<TT/>} />
            </PieChart>
          </ResponsiveContainer>
          <div style={{ display:'flex', flexDirection:'column', gap:6, marginTop:4 }}>
            {typeData.map((e,i) => (
              <div key={i} style={{ display:'flex', justifyContent:'space-between', alignItems:'center' }}>
                <span style={{ display:'flex', alignItems:'center', gap:6, fontSize:11, color:C.muted }}>
                  <span style={{ width:8,height:8,borderRadius:2,background:TYPE_C[e.name]||C.purple,display:'inline-block' }}/>
                  {e.name}
                </span>
                <span style={{ fontSize:11, color:C.text }}>{pct(e.value,a.totalGeral)}</span>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div style={s.card}>
        <p style={{ ...s.label, marginBottom:16 }}>Despesas por Categoria (Anual)</p>
        <ResponsiveContainer width="100%" height={Math.max(200,catData.length*36)}>
          <BarChart data={catData} layout="vertical" margin={{top:0,right:80,bottom:0,left:90}}>
            <XAxis type="number" tick={{fill:C.muted,fontSize:11}} axisLine={false} tickLine={false} tickFormatter={v=>hide?'•••':`${(v/1000).toFixed(0)}k`} />
            <YAxis type="category" dataKey="name" tick={{fill:C.muted,fontSize:12}} axisLine={false} tickLine={false} width={90} />
            <Tooltip content={<TT/>} />
            <Bar dataKey="value" radius={[0,4,4,0]} label={{ position:'right', fill:C.muted, fontSize:11, formatter:v=>maskFmtK(v, hide) }}>
              {catData.map((e,i) => <Cell key={i} fill={e.fill||C.muted} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

// ─── TAB: ENCONTRO MENSAL ───
function EncontroTab({ a, viewMode, selectedMonth, currentMonth }) {
  const hide = useHideBalances();
  const { paidMonths } = usePaidMonths();

  if (viewMode === 'monthly') {
    const mData = a.monthly[selectedMonth];
    const isPaid = !!paidMonths[selectedMonth];
    if (mData.total === 0) return (
      <div style={{ ...s.card, textAlign:'center', padding:48 }}>
        <p style={{ fontSize:32, marginBottom:12 }}>📭</p>
        <p style={{ color:C.muted }}>Sem lançamentos em {mData.mes}.</p>
      </div>
    );

    const dirColor = mData.direction==='g2b'?C.red:mData.direction==='b2g'?C.green:C.muted;
    const dirText = mData.direction==='g2b'?'Guilherme deve transferir para Bruna':mData.direction==='b2g'?'Bruna deve transferir para Guilherme':'Zerado';
    const effectiveG = isPaid ? 0 : mData.gOwesB;
    const effectiveB = isPaid ? 0 : mData.bOwesG;
    const effectiveNet = isPaid ? 0 : mData.net;

    return (
      <div>
        <MonthlySummaryBanner mData={mData} rows={a.rows} currentMonth={currentMonth} />

        <div style={{ ...s.card, background:`linear-gradient(135deg, ${C.card}, #0a2040)`, border:`1px solid ${isPaid ? C.green : C.gold}55`, marginBottom:20 }}>
          <p style={{ ...s.label, color:isPaid?C.green:C.gold, marginBottom:12 }}>
            {isPaid ? '✓ Encontro de Contas — Liquidado' : '⚖️ Encontro de Contas'} — {mData.mes} {YEAR}
          </p>
          <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(180px,1fr))', gap:16 }}>
            <div style={{ background:C.redDim, border:`1px solid ${C.red}40`, borderRadius:10, padding:'14px 16px' }}>
              <p style={{ ...s.muted, fontSize:11, marginBottom:4 }}>Guilherme a pagar (à Bruna)</p>
              <p style={{ fontSize:22, fontFamily:'monospace', color:C.red, fontWeight:700 }}>{maskFmt(effectiveG, hide)}</p>
            </div>
            <div style={{ background:C.greenDim, border:`1px solid ${C.green}40`, borderRadius:10, padding:'14px 16px' }}>
              <p style={{ ...s.muted, fontSize:11, marginBottom:4 }}>Bruna a pagar (ao Guilherme)</p>
              <p style={{ fontSize:22, fontFamily:'monospace', color:C.green, fontWeight:700 }}>{maskFmt(effectiveB, hide)}</p>
            </div>
            <div style={{ background:`${isPaid?C.green:dirColor}15`, border:`1px solid ${isPaid?C.green:dirColor}40`, borderRadius:10, padding:'14px 16px' }}>
              <p style={{ ...s.muted, fontSize:11, marginBottom:4 }}>💳 Transferência Líquida</p>
              <p style={{ fontSize:26, fontFamily:'monospace', color:isPaid?C.green:dirColor, fontWeight:700 }}>{maskFmt(Math.abs(effectiveNet), hide)}</p>
              <p style={{ fontSize:11, color:isPaid?C.green:dirColor, marginTop:4, fontWeight:600 }}>{isPaid ? '✓ Liquidado' : dirText}</p>
            </div>
          </div>
        </div>

        <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(180px,1fr))', gap:12, marginBottom:20 }}>
          <div style={s.cardSm}>
            <p style={s.label}>Total no Mês</p>
            <p style={{ fontSize:20, fontFamily:'monospace', color:C.text, marginTop:4 }}>{maskFmt(mData.total, hide)}</p>
          </div>
          <div style={s.cardSm}>
            <p style={s.label}>Guilherme pagou</p>
            <p style={{ fontSize:20, fontFamily:'monospace', color:C.blue, marginTop:4 }}>{maskFmt(mData.Guilherme, hide)}</p>
            <p style={{ fontSize:11, color:C.muted, marginTop:2 }}>{pct(mData.Guilherme,mData.total)} do total</p>
          </div>
          <div style={s.cardSm}>
            <p style={s.label}>Bruna pagou</p>
            <p style={{ fontSize:20, fontFamily:'monospace', color:C.amber, marginTop:4 }}>{maskFmt(mData.Bruna, hide)}</p>
            <p style={{ fontSize:11, color:C.muted, marginTop:2 }}>{pct(mData.Bruna,mData.total)} do total</p>
          </div>
        </div>

        {mData.topCats?.length > 0 && (
          <div style={s.card}>
            <p style={{ ...s.label, marginBottom:12 }}>Top Categorias</p>
            {mData.topCats.map(([cat,val]) => (
              <div key={cat} style={{ display:'flex', justifyContent:'space-between', alignItems:'center', padding:'8px 0', borderBottom:`1px solid ${C.border}` }}>
                <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                  <div style={{ width:3, height:32, borderRadius:2, background:CAT_C[cat]||C.muted }}/>
                  <div>
                    <p style={{ fontSize:13, color:C.text }}>{cat}</p>
                    <p style={{ fontSize:11, color:C.muted }}>{pct(val,mData.total)}</p>
                  </div>
                </div>
                <p style={{ fontSize:14, fontFamily:'monospace', color:CAT_C[cat]||C.text }}>{maskFmt(val, hide)}</p>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  // Visão anual: total a liquidar exclui meses já liquidados
  const pending = pendingSettlement(a.monthly, paidMonths);
  const months = a.monthly.filter(m => m.gOwesB > 0 || m.bOwesG > 0);

  return (
    <div>
      <div style={{ ...s.card, marginBottom:20, background:`linear-gradient(135deg, ${C.card}, ${C.panel})`, border:`1px solid ${C.gold}55` }}>
        <p style={{ ...s.label, color:C.gold, marginBottom:12 }}>⚖️ Encontro de Contas — Resumo Anual {YEAR}</p>
        <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(180px,1fr))', gap:20 }}>
          <div>
            <p style={s.muted}>Guilherme a pagar (pendente)</p>
            <p style={{ fontSize:24, fontFamily:'monospace', color:C.red, fontWeight:700, marginTop:4 }}>{maskFmt(pending.g2b, hide)}</p>
          </div>
          <div>
            <p style={s.muted}>Bruna a pagar (pendente)</p>
            <p style={{ fontSize:24, fontFamily:'monospace', color:C.green, fontWeight:700, marginTop:4 }}>{maskFmt(pending.b2g, hide)}</p>
          </div>
          <div style={{ borderLeft:`1px solid ${C.border}`, paddingLeft:20 }}>
            <p style={s.muted}>Liquidação Líquida</p>
            <p style={{ fontSize:28, fontFamily:'monospace', color:pending.net>0?C.red:pending.net<0?C.green:C.muted, fontWeight:700, marginTop:4 }}>{maskFmt(Math.abs(pending.net), hide)}</p>
            <p style={{ fontSize:12, color:pending.net>0?C.red:C.green, fontWeight:600, marginTop:4 }}>
              {pending.net>0.01?'🔴 Guilherme transfere para Bruna':pending.net<-0.01?'🟢 Bruna transfere para Guilherme':'✅ Zerado'}
            </p>
          </div>
        </div>
        <p style={{ ...s.muted, marginTop:12, fontSize:11 }}>
          Meses já liquidados não compõem este saldo.
        </p>
      </div>

      <p style={{ ...s.label, marginBottom:12 }}>📅 Mês a Mês</p>
      <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(300px,1fr))', gap:12 }}>
        {months.map(m => {
          const isPaid = !!paidMonths[m.idx];
          const color = isPaid ? C.green : m.direction==='g2b'?C.red:m.direction==='b2g'?C.green:C.muted;
          const txt = isPaid ? '✓ Liquidado' : m.direction==='g2b'?'Guilherme → Bruna':m.direction==='b2g'?'Bruna → Guilherme':'Zerado';
          const eff = isPaid ? 0 : Math.abs(m.net);
          return (
            <div key={m.idx} style={{ ...s.card, borderLeft:`4px solid ${color}` }}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:10 }}>
                <p style={{ fontSize:16, fontWeight:700, color:C.text }}>{m.mes}</p>
                <Badge text={txt} color={color} />
              </div>
              <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:8, marginBottom:10, fontSize:12 }}>
                <div style={{ background:C.panel, padding:'6px 10px', borderRadius:6 }}>
                  <p style={{ color:C.muted, fontSize:10 }}>G a pagar</p>
                  <p style={{ color:isPaid?C.muted:C.red, fontFamily:'monospace', marginTop:2, textDecoration:isPaid?'line-through':'none' }}>{maskFmt(m.gOwesB, hide)}</p>
                </div>
                <div style={{ background:C.panel, padding:'6px 10px', borderRadius:6 }}>
                  <p style={{ color:C.muted, fontSize:10 }}>B a pagar</p>
                  <p style={{ color:isPaid?C.muted:C.green, fontFamily:'monospace', marginTop:2, textDecoration:isPaid?'line-through':'none' }}>{maskFmt(m.bOwesG, hide)}</p>
                </div>
              </div>
              <div style={{ borderTop:`1px dashed ${C.border}`, paddingTop:8, display:'flex', justifyContent:'space-between', alignItems:'center' }}>
                <span style={{ fontSize:11, color:C.muted }}>Transferência líquida</span>
                <span style={{ fontSize:16, fontFamily:'monospace', fontWeight:700, color }}>{maskFmt(eff, hide)}</span>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── TAB: SPLIT ANUAL ───
function SplitTab({ a, viewMode, selectedMonth, currentMonth }) {
  const hide = useHideBalances();
  const { paidMonths } = usePaidMonths();

  if (viewMode === 'monthly') {
    const m = a.monthly[selectedMonth];
    const isPaid = !!paidMonths[selectedMonth];
    const g2bTotal = isPaid ? 0 : m.gOwesB;
    const b2gTotal = isPaid ? 0 : m.bOwesG;
    const netMonth = isPaid ? 0 : m.net;
    const g2bRows = a.rows
      .filter(r => r.q === 'Guilherme' && r.p === 'Bruna')
      .map(r => {
        const v = r.m[selectedMonth] || 0;
        return v > 0 ? { ...r, total: v } : null;
      })
      .filter(Boolean);
    const b2gRows = a.rows
      .filter(r => r.q === 'Bruna' && r.p === 'Guilherme')
      .map(r => {
        const v = r.m[selectedMonth] || 0;
        return v > 0 ? { ...r, total: v } : null;
      })
      .filter(Boolean);
    const monthFixed = a.rows.filter(r => r.c === 'Despesa Fixa').reduce((s, r) => s + (r.m[selectedMonth] || 0), 0);
    const monthTotal = m.total;
    const monthVar = monthTotal - monthFixed;
    return (
      <div>
        <MonthlySummaryBanner mData={m} rows={a.rows} currentMonth={currentMonth} />

        <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))', gap:16, marginBottom:24 }}>
          <div style={{ ...s.card, borderTop:`3px solid ${C.blue}` }}>
            <p style={s.label}>Guilherme a pagar à Bruna</p>
            <p style={{ ...s.val, color:isPaid?C.green:C.red, marginTop:6 }}>{maskFmt(g2bTotal, hide)}</p>
            <p style={s.muted}>{g2bRows.length} lançamentos {isPaid && '(liquidado)'}</p>
          </div>
          <div style={{ ...s.card, borderTop:`3px solid ${C.amber}` }}>
            <p style={s.label}>Bruna a pagar ao Guilherme</p>
            <p style={{ ...s.val, color:isPaid?C.green:C.green, marginTop:6 }}>{maskFmt(b2gTotal, hide)}</p>
            <p style={s.muted}>{b2gRows.length} lançamentos {isPaid && '(liquidado)'}</p>
          </div>
          <div style={{ ...s.card, borderTop:`3px solid ${isPaid?C.green:netMonth>0?C.red:C.green}` }}>
            <p style={s.label}>Saldo Líquido</p>
            <p style={{ ...s.val, color:isPaid?C.green:netMonth>0?C.red:C.green, marginTop:6 }}>{maskFmt(Math.abs(netMonth), hide)}</p>
            <p style={s.muted}>{isPaid ? '✓ Liquidado' : netMonth>0.01?'Guilherme → Bruna':netMonth<-0.01?'Bruna → Guilherme':'Zerado'}</p>
          </div>
        </div>
        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:16, marginBottom:16 }}>
          <div style={s.card}>
            <p style={{ ...s.label, color:C.red, marginBottom:12 }}>Guilherme → Bruna</p>
            <div style={{ display:'flex', flexDirection:'column', gap:8, maxHeight:400, overflowY:'auto' }}>
              {g2bRows.sort((a,b)=>b.total-a.total).map((r,i) => (
                <div key={i} style={{ display:'flex', justifyContent:'space-between', padding:'8px 12px', background:C.panel, borderRadius:8, borderLeft:`3px solid ${CAT_C[r.c]||C.muted}` }}>
                  <div style={{ minWidth:0 }}>
                    <p style={{ fontSize:13, color:C.text, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.d}</p>
                    <p style={{ fontSize:11, color:C.muted }}>{r.c} • {r.t}</p>
                  </div>
                    <p style={{ fontSize:13, fontFamily:'monospace', color:C.red, alignSelf:'center', marginLeft:8 }}>{maskFmt(r.total, hide)}</p>
                </div>
              ))}
            </div>
          </div>
          <div style={s.card}>
            <p style={{ ...s.label, color:C.green, marginBottom:12 }}>Bruna → Guilherme</p>
            {b2gRows.length > 0 ? (
              <div style={{ display:'flex', flexDirection:'column', gap:8, maxHeight:400, overflowY:'auto' }}>
                {b2gRows.sort((a,b)=>b.total-a.total).map((r,i) => (
                  <div key={i} style={{ display:'flex', justifyContent:'space-between', padding:'8px 12px', background:C.panel, borderRadius:8, borderLeft:`3px solid ${CAT_C[r.c]||C.muted}` }}>
                    <div style={{ minWidth:0 }}>
                      <p style={{ fontSize:13, color:C.text, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.d}</p>
                      <p style={{ fontSize:11, color:C.muted }}>{r.c} • {r.t}</p>
                    </div>
                    <p style={{ fontSize:13, fontFamily:'monospace', color:C.green, alignSelf:'center', marginLeft:8 }}>{maskFmt(r.total, hide)}</p>
                  </div>
                ))}
              </div>
            ) : <p style={s.muted}>Nenhum repasse nesta direção.</p>}
          </div>
        </div>
        <div style={s.card}>
          <p style={{ ...s.label, marginBottom:12 }}>Fixo vs Variável (Mês)</p>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:16 }}>
            <div>
              <p style={{ ...s.muted, marginBottom:6 }}>Despesas Fixas</p>
              <p style={{ fontSize:22, fontFamily:'monospace', color:C.red }}>{maskFmt(monthFixed, hide)}</p>
              <p style={s.muted}>{pct(monthFixed, monthTotal)} do total</p>
            </div>
            <div>
              <p style={{ ...s.muted, marginBottom:6 }}>Despesas Variáveis</p>
              <p style={{ fontSize:22, fontFamily:'monospace', color:C.blue }}>{maskFmt(monthVar, hide)}</p>
              <p style={s.muted}>{pct(monthVar, monthTotal)} do total</p>
            </div>
          </div>
          <div style={{ marginTop:16, height:8, borderRadius:4, background:C.dim, overflow:'hidden' }}>
            <div style={{ height:'100%', width:pct(monthFixed, monthTotal), background:C.red, borderRadius:4 }}/>
          </div>
        </div>
      </div>
    );
  }

  const pending = pendingSettlement(a.monthly, paidMonths);
  const splitRows = a.rows.filter(r => r.q !== r.p && r.total > 0);
  const g2bRows = splitRows.filter(r => r.q==='Guilherme'&&r.p==='Bruna');
  const b2gRows = splitRows.filter(r => r.q==='Bruna'&&r.p==='Guilherme');
  return (
    <div>
      <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))', gap:16, marginBottom:24 }}>
        <div style={{ ...s.card, borderTop:`3px solid ${C.blue}` }}>
          <p style={s.label}>Guilherme a pagar (pendente)</p>
          <p style={{ ...s.val, color:C.red, marginTop:6 }}>{maskFmt(pending.g2b, hide)}</p>
          <p style={s.muted}>{g2bRows.length} lançamentos no ano</p>
        </div>
        <div style={{ ...s.card, borderTop:`3px solid ${C.amber}` }}>
          <p style={s.label}>Bruna a pagar (pendente)</p>
          <p style={{ ...s.val, color:C.green, marginTop:6 }}>{maskFmt(pending.b2g, hide)}</p>
          <p style={s.muted}>{b2gRows.length} lançamentos no ano</p>
        </div>
        <div style={{ ...s.card, borderTop:`3px solid ${pending.net>0?C.red:C.green}` }}>
          <p style={s.label}>Saldo Líquido</p>
          <p style={{ ...s.val, color:pending.net>0?C.red:C.green, marginTop:6 }}>{maskFmt(Math.abs(pending.net), hide)}</p>
          <p style={s.muted}>{pending.net>0.01?'Guilherme → Bruna':pending.net<-0.01?'Bruna → Guilherme':'Zerado'}</p>
        </div>
      </div>
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:16, marginBottom:16 }}>
        <div style={s.card}>
          <p style={{ ...s.label, color:C.red, marginBottom:12 }}>Guilherme → Bruna</p>
          <div style={{ display:'flex', flexDirection:'column', gap:8, maxHeight:400, overflowY:'auto' }}>
            {g2bRows.sort((a,b)=>b.total-a.total).map((r,i) => (
              <div key={i} style={{ display:'flex', justifyContent:'space-between', padding:'8px 12px', background:C.panel, borderRadius:8, borderLeft:`3px solid ${CAT_C[r.c]||C.muted}` }}>
                <div style={{ minWidth:0 }}>
                  <p style={{ fontSize:13, color:C.text, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.d}</p>
                  <p style={{ fontSize:11, color:C.muted }}>{r.c} • {r.t}</p>
                </div>
                <p style={{ fontSize:13, fontFamily:'monospace', color:C.red, alignSelf:'center', marginLeft:8 }}>{maskFmt(r.total, hide)}</p>
              </div>
            ))}
          </div>
        </div>
        <div style={s.card}>
          <p style={{ ...s.label, color:C.green, marginBottom:12 }}>Bruna → Guilherme</p>
          {b2gRows.length > 0 ? (
            <div style={{ display:'flex', flexDirection:'column', gap:8, maxHeight:400, overflowY:'auto' }}>
              {b2gRows.sort((a,b)=>b.total-a.total).map((r,i) => (
                <div key={i} style={{ display:'flex', justifyContent:'space-between', padding:'8px 12px', background:C.panel, borderRadius:8, borderLeft:`3px solid ${CAT_C[r.c]||C.muted}` }}>
                  <div style={{ minWidth:0 }}>
                    <p style={{ fontSize:13, color:C.text, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.d}</p>
                    <p style={{ fontSize:11, color:C.muted }}>{r.c} • {r.t}</p>
                  </div>
                  <p style={{ fontSize:13, fontFamily:'monospace', color:C.green, alignSelf:'center', marginLeft:8 }}>{maskFmt(r.total, hide)}</p>
                </div>
              ))}
            </div>
          ) : <p style={s.muted}>Nenhum repasse nesta direção.</p>}
        </div>
      </div>
      <div style={s.card}>
        <p style={{ ...s.label, marginBottom:12 }}>Fixo vs Variável</p>
        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:16 }}>
          <div>
            <p style={{ ...s.muted, marginBottom:6 }}>Despesas Fixas</p>
            <p style={{ fontSize:22, fontFamily:'monospace', color:C.red }}>{maskFmt(a.fixedTotal, hide)}</p>
            <p style={s.muted}>{pct(a.fixedTotal,a.totalGeral)} do total</p>
          </div>
          <div>
            <p style={{ ...s.muted, marginBottom:6 }}>Despesas Variáveis</p>
            <p style={{ fontSize:22, fontFamily:'monospace', color:C.blue }}>{maskFmt(a.varTotal, hide)}</p>
            <p style={s.muted}>{pct(a.varTotal,a.totalGeral)} do total</p>
          </div>
        </div>
        <div style={{ marginTop:16, height:8, borderRadius:4, background:C.dim, overflow:'hidden' }}>
          <div style={{ height:'100%', width:pct(a.fixedTotal,a.totalGeral), background:C.red, borderRadius:4 }}/>
        </div>
      </div>
    </div>
  );
}

// ─── TAB: TENDÊNCIAS ───
function TendenciasTab({ a }) {
  const hide = useHideBalances();
  const momData = a.monthlyActive.map((m,i,arr) => {
    const prev = arr[i-1];
    const mom = prev&&prev.total ? ((m.total-prev.total)/prev.total*100) : 0;
    return { ...m, mom: Math.round(mom*10)/10 };
  });

  return (
    <div>
      <div style={{ ...s.card, marginBottom:16 }}>
        <p style={{ ...s.label, marginBottom:16 }}>Evolução Mensal (R$)</p>
        <ResponsiveContainer width="100%" height={260}>
          <AreaChart data={a.monthlyActive} margin={{top:0,right:10,bottom:0,left:20}}>
            <defs>
              <linearGradient id="gG" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor={C.blue} stopOpacity={0.3}/>
                <stop offset="95%" stopColor={C.blue} stopOpacity={0}/>
              </linearGradient>
              <linearGradient id="gB" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor={C.amber} stopOpacity={0.3}/>
                <stop offset="95%" stopColor={C.amber} stopOpacity={0}/>
              </linearGradient>
            </defs>
            <XAxis dataKey="mes" tick={{fill:C.muted,fontSize:11}} axisLine={false} tickLine={false} />
            <YAxis tick={{fill:C.muted,fontSize:11}} axisLine={false} tickLine={false} tickFormatter={v=>hide?'•••':`${(v/1000).toFixed(0)}k`} />
            <Tooltip content={<TT/>} />
            <Area type="monotone" dataKey="Guilherme" stroke={C.blue} fill="url(#gG)" strokeWidth={2} />
            <Area type="monotone" dataKey="Bruna" stroke={C.amber} fill="url(#gB)" strokeWidth={2} />
          </AreaChart>
        </ResponsiveContainer>
      </div>
      <div style={s.card}>
        <p style={{ ...s.label, marginBottom:16 }}>Variação Mensal (MoM %)</p>
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={momData.slice(1)} margin={{top:10,right:10,bottom:0,left:10}}>
            <XAxis dataKey="mes" tick={{fill:C.muted,fontSize:11}} axisLine={false} tickLine={false} />
            <YAxis tick={{fill:C.muted,fontSize:11}} axisLine={false} tickLine={false} tickFormatter={v=>`${v}%`} />
            <Tooltip formatter={v=>`${v}%`} contentStyle={{background:C.panel,border:`1px solid ${C.border}`,borderRadius:8}} />
            <Bar dataKey="mom" radius={4}>
              {momData.slice(1).map((e,i)=><Cell key={i} fill={e.mom>=0?C.red:C.green} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(160px,1fr))', gap:16, marginTop:16 }}>
        <div style={s.cardSm}>
          <p style={s.label}>Pico de Gasto</p>
          <p style={{ fontSize:18, fontFamily:'monospace', color:C.red, marginTop:4 }}>{a.monthlyActive.reduce((p,c)=>c.total>p.total?c:p,{total:0,mes:'-'}).mes}</p>
          <p style={s.muted}>{maskFmt(a.monthlyActive.reduce((p,c)=>c.total>p.total?c:p,{total:0}).total, hide)}</p>
        </div>
        <div style={s.cardSm}>
          <p style={s.label}>Média Mensal</p>
          <p style={{ fontSize:18, fontFamily:'monospace', color:C.gold, marginTop:4 }}>{maskFmt(a.totalGeral/Math.max(1,a.monthlyActive.length), hide)}</p>
          <p style={s.muted}>{a.monthlyActive.length} meses ativos</p>
        </div>
        <div style={s.cardSm}>
          <p style={s.label}>Menor Mês</p>
          <p style={{ fontSize:18, fontFamily:'monospace', color:C.green, marginTop:4 }}>{a.monthlyActive.reduce((p,c)=>c.total<p.total?c:p,{total:Infinity,mes:'-'}).mes}</p>
          <p style={s.muted}>{maskFmt(a.monthlyActive.reduce((p,c)=>c.total<p.total?c:p,{total:Infinity}).total, hide)}</p>
        </div>
      </div>
    </div>
  );
}

// ─── TAB: CATEGORIAS ───
function CategoriasTab({ a, viewMode, selectedMonth }) {
  const hide = useHideBalances();
  if (viewMode === 'monthly') {
    const monthRows = a.rows
      .map(r => ({ ...r, val: r.m[selectedMonth] || 0 }))
      .filter(r => r.val > 0);

    const byCat = {};
    const byDesc = {};

    monthRows.forEach(r => {
      byCat[r.c] = (byCat[r.c] || 0) + r.val;
      byDesc[r.d] = (byDesc[r.d] || 0) + r.val;
    });

    const catEntries = Object.entries(byCat).sort((x,y)=>y[1]-x[1]);
    const topDesc = Object.entries(byDesc).sort((x,y)=>y[1]-x[1]).slice(0,10);
    const monthTotal = catEntries.reduce((s,[,v])=>s+v,0);

    if (monthRows.length === 0) {
      return (
        <div style={{ ...s.card, textAlign:'center', padding:48 }}>
          <p style={{ fontSize:32, marginBottom:12 }}>📭</p>
          <p style={{ color:C.muted }}>Sem categorias com lançamentos em {ML[selectedMonth]}.</p>
        </div>
      );
    }

    return (
      <div>
        <div style={{ ...s.card, marginBottom:20, border:`1px solid ${C.gold}40`, background:`linear-gradient(135deg, ${C.card}, #0A1E35)` }}>
          <p style={{ ...s.label, color:C.gold }}>🏷 Categorias — {ML[selectedMonth]} {YEAR}</p>
          <p style={{ fontSize:28, fontWeight:700, marginTop:10, color:C.text, fontFamily:'monospace' }}>
            {maskFmt(monthTotal, hide)}
          </p>
          <p style={{ ...s.muted, marginTop:6 }}>
            {catEntries.length} categorias ativas • {monthRows.length} lançamentos no mês
          </p>
        </div>

        <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(220px,1fr))', gap:12, marginBottom:20 }}>
          {catEntries.map(([cat,val]) => (
            <div key={cat} style={{ ...s.cardSm, borderLeft:`4px solid ${CAT_C[cat] || C.muted}` }}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start' }}>
                <div>
                  <p style={{ ...s.label, color:CAT_C[cat] || C.muted }}>{cat}</p>
                  <p style={{ fontSize:20, fontFamily:'monospace', color:C.text, marginTop:4 }}>{maskFmt(val, hide)}</p>
                </div>
                <span style={{ fontSize:12, color:CAT_C[cat] || C.muted, fontWeight:700 }}>{pct(val,monthTotal)}</span>
              </div>
              <div style={{ marginTop:10, height:5, borderRadius:4, background:C.dim, overflow:'hidden' }}>
                <div style={{ height:'100%', width:pct(val,monthTotal), background:CAT_C[cat] || C.muted, borderRadius:4 }}/>
              </div>
            </div>
          ))}
        </div>

        <div style={s.card}>
          <p style={{ ...s.label, marginBottom:16 }}>Top 10 Maiores Despesas do Mês</p>
          <div style={{ display:'flex', flexDirection:'column', gap:8 }}>
            {topDesc.map(([desc,val],i) => {
              const matched = monthRows.find(r=>r.d===desc);
              const cat = matched?.c || 'Outros';
              return (
                <div key={i} style={{ display:'flex', justifyContent:'space-between', alignItems:'center', padding:'8px 12px', background:C.panel, borderRadius:8 }}>
                  <div style={{ display:'flex', alignItems:'center', gap:12, minWidth:0 }}>
                    <span style={{ fontSize:11, color:C.muted, fontFamily:'monospace', minWidth:20 }}>#{i+1}</span>
                    <div style={{ minWidth:0 }}>
                      <p style={{ fontSize:13, color:C.text, whiteSpace:'nowrap', overflow:'hidden', textOverflow:'ellipsis' }}>{desc}</p>
                      <span style={s.tag(CAT_C[cat]||C.muted)}>{cat}</span>
                    </div>
                  </div>
                  <div style={{ textAlign:'right', marginLeft:12 }}>
                    <p style={{ fontSize:14, fontFamily:'monospace', color:C.text }}>{maskFmt(val, hide)}</p>
                    <p style={{ fontSize:11, color:C.muted }}>{pct(val,monthTotal)}</p>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    );
  }

  const catEntries = Object.entries(a.byCat).sort((x,y)=>y[1]-x[1]);

  return (
    <div>
      <div style={{ ...s.card, marginBottom:20, border:`1px solid ${C.gold}40`, background:`linear-gradient(135deg, ${C.card}, #0A1E35)` }}>
        <p style={{ ...s.label, color:C.gold }}>🏷 Categorias — Visão Anual {YEAR}</p>
        <p style={{ fontSize:28, fontWeight:700, marginTop:10, color:C.text, fontFamily:'monospace' }}>
          {maskFmt(a.totalGeral, hide)}
        </p>
        <p style={{ ...s.muted, marginTop:6 }}>{catEntries.length} categorias ativas no ano</p>
      </div>

      <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(220px,1fr))', gap:12, marginBottom:20 }}>
        {catEntries.map(([cat,val]) => (
          <div key={cat} style={{ ...s.cardSm, borderLeft:`4px solid ${CAT_C[cat] || C.muted}` }}>
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start' }}>
              <div>
                <p style={{ ...s.label, color:CAT_C[cat] || C.muted }}>{cat}</p>
                <p style={{ fontSize:20, fontFamily:'monospace', color:C.text, marginTop:4 }}>{maskFmt(val, hide)}</p>
              </div>
              <span style={{ fontSize:12, color:CAT_C[cat] || C.muted, fontWeight:700 }}>{pct(val,a.totalGeral)}</span>
            </div>
            <div style={{ marginTop:10, height:4, borderRadius:2, background:C.dim }}>
              <div style={{ height:'100%', width:pct(val,a.totalGeral), background:CAT_C[cat] || C.muted, borderRadius:2 }}/>
            </div>
          </div>
        ))}
      </div>

      <div style={s.card}>
        <p style={{ ...s.label, marginBottom:16 }}>Top 10 Maiores Despesas (Item)</p>
        <div style={{ display:'flex', flexDirection:'column', gap:8 }}>
          {a.topDesc.map(([desc,val],i) => {
            const matched = a.rows.find(r=>r.d===desc);
            const cat = matched?.c||'Outros';
            return (
              <div key={i} style={{ display:'flex', justifyContent:'space-between', alignItems:'center', padding:'8px 12px', background:C.panel, borderRadius:8 }}>
                <div style={{ display:'flex', alignItems:'center', gap:12, minWidth:0 }}>
                  <span style={{ fontSize:11, color:C.muted, fontFamily:'monospace', minWidth:20 }}>#{i+1}</span>
                  <div style={{ minWidth:0 }}>
                    <p style={{ fontSize:13, color:C.text, whiteSpace:'nowrap', overflow:'hidden', textOverflow:'ellipsis' }}>{desc}</p>
                    <span style={s.tag(CAT_C[cat]||C.muted)}>{cat}</span>
                  </div>
                </div>
                <div style={{ textAlign:'right', marginLeft:12 }}>
                  <p style={{ fontSize:14, fontFamily:'monospace', color:C.text }}>{maskFmt(val, hide)}</p>
                  <p style={{ fontSize:11, color:C.muted }}>{pct(val,a.totalGeral)}</p>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ─── TAB: ADICIONAR ───
function AdicionarTab({ rows, onChange }) {
  const [form, setForm] = useState({ t:'Cartão de Crédito',q:'Bruna',p:'Bruna',c:'Mercado',d:'',mode:'unico',mes:0,valor:'',mesInicial:0,parcelas:1,valorParcela:'' });
  const [done, setDone] = useState('');
  const update = (k,v) => setForm(f=>({...f,[k]:v}));

  const submit = () => {
    if (!form.d.trim()) { setDone('❌ Descrição obrigatória'); return; }
    const m = Array(12).fill(0);
    if (form.mode==='unico') {
      const v = parseFloat(String(form.valor).replace(',','.'));
      if (!v||v<=0) { setDone('❌ Valor inválido'); return; }
      m[form.mes] = v;
    } else {
      const v = parseFloat(String(form.valorParcela).replace(',','.'));
      const n = parseInt(form.parcelas);
      if (!v||v<=0||!n||n<1) { setDone('❌ Valores inválidos'); return; }
      for (let i=0;i<n&&(form.mesInicial+i)<12;i++) m[form.mesInicial+i]=v;
    }
    onChange([...rows, { id:uid(), t:form.t, q:form.q, p:form.p, c:form.c, d:form.d.trim(), m, total:m.reduce((s,v)=>s+v,0) }]);
    setDone(`✅ "${form.d}" adicionada!`);
    setForm(f=>({...f,d:'',valor:'',valorParcela:'',parcelas:1}));
    setTimeout(()=>setDone(''), 4000);
  };

  return (
    <div>
      <div style={{ ...s.card, marginBottom:16, borderColor:C.gold }}>
        <p style={{ ...s.label, color:C.gold, marginBottom:6 }}>📝 Nova Despesa</p>
        <p style={{ ...s.muted, lineHeight:1.6 }}>Para despesas <strong style={{color:C.text}}>compartilhadas</strong>, cadastre duas linhas (uma por pessoa).</p>
      </div>
      <div style={s.card}>
        <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))', gap:16 }}>
          <Field label="Tipo"><select style={s.input} value={form.t} onChange={e=>update('t',e.target.value)}>{TIPOS.map(t=><option key={t} value={t}>{t}</option>)}</select></Field>
          <Field label="Categoria"><select style={s.input} value={form.c} onChange={e=>update('c',e.target.value)}>{CATEGORIAS.map(c=><option key={c} value={c}>{c}</option>)}</select></Field>
          <Field label="Quem pagou"><select style={s.input} value={form.q} onChange={e=>update('q',e.target.value)}>{PESSOAS.map(p=><option key={p} value={p}>{p}</option>)}</select></Field>
          <Field label="Para quem"><select style={s.input} value={form.p} onChange={e=>update('p',e.target.value)}>{PESSOAS.map(p=><option key={p} value={p}>{p}</option>)}</select></Field>
          <Field label="Descrição" span={2}><input style={s.input} value={form.d} onChange={e=>update('d',e.target.value)} placeholder="ex: Conta de luz, Mercado..." /></Field>
        </div>
        <div style={{ marginTop:24, paddingTop:20, borderTop:`1px dashed ${C.border}` }}>
          <p style={{ ...s.label, marginBottom:12 }}>Modo de Pagamento</p>
          <div style={{ display:'flex', gap:8, marginBottom:16 }}>
            {[{k:'unico',l:'💵 À Vista'},{k:'parcelado',l:'💳 Parcelado'}].map(opt=>(
              <button key={opt.k} onClick={()=>update('mode',opt.k)}
                style={{ flex:1, padding:'10px 16px', borderRadius:8, border:`1px solid ${form.mode===opt.k?C.gold:C.border}`, background:form.mode===opt.k?C.goldDim:'transparent', color:form.mode===opt.k?C.gold:C.muted, fontSize:13, fontWeight:600, cursor:'pointer', transition:'all 0.2s' }}>
                {opt.l}
              </button>
            ))}
          </div>
          {form.mode==='unico' ? (
            <div style={{ display:'grid', gridTemplateColumns:'1fr 2fr', gap:16 }}>
              <Field label="Mês"><select style={s.input} value={form.mes} onChange={e=>update('mes',parseInt(e.target.value))}>{ML.map((m,i)=><option key={i} value={i}>{m}</option>)}</select></Field>
              <Field label="Valor (R$)"><input style={s.input} value={form.valor} onChange={e=>update('valor',e.target.value)} placeholder="0,00" /></Field>
            </div>
          ) : (
            <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(150px,1fr))', gap:16 }}>
              <Field label="Mês inicial"><select style={s.input} value={form.mesInicial} onChange={e=>update('mesInicial',parseInt(e.target.value))}>{ML.map((m,i)=><option key={i} value={i}>{m}</option>)}</select></Field>
              <Field label="Nº parcelas"><input style={s.input} type="number" min="1" max="12" value={form.parcelas} onChange={e=>update('parcelas',e.target.value)} /></Field>
              <Field label="Valor por parcela"><input style={s.input} value={form.valorParcela} onChange={e=>update('valorParcela',e.target.value)} placeholder="0,00" /></Field>
            </div>
          )}
        </div>
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginTop:24 }}>
          <p style={{ fontSize:12, color:done.startsWith('✅')?C.green:done.startsWith('❌')?C.red:C.muted }}>{done}</p>
          <button onClick={submit} style={s.btn(C.gold)}>✨ Adicionar</button>
        </div>
      </div>
    </div>
  );
}

// ─── TAB: TABELA ───
// Filtros: pessoa, tipo, categoria, mês inicial/final, status (todos|liquidado|a pagar), busca por descrição.
// Os filtros só são aplicados após o usuário clicar em "Aplicar filtros" (batch filtering).
function TabelaTab({ rows, onChange }) {
  const { paidMonths } = usePaidMonths();
  const hide = useHideBalances();

  // Rascunho de filtros (o usuário ajusta antes de aplicar).
  const initial = { search:'', q:'all', t:'all', c:'all', monthFrom:0, monthTo:11, status:'all' };
  const [draft, setDraft] = useState(initial);
  const [applied, setApplied] = useState(initial);
  const [editId, setEditId] = useState(null);
  const [editForm, setEditForm] = useState(null);

  const dirty = JSON.stringify(draft) !== JSON.stringify(applied);

  const getWindowSum = (row, f) => {
    const lo = Math.min(f.monthFrom, f.monthTo);
    const hi = Math.max(f.monthFrom, f.monthTo);
    let sumWindow = 0;
    for (let i=lo;i<=hi;i++) sumWindow += row.m[i] || 0;
    return sumWindow;
  };

  const filtered = useMemo(() => rows.filter(r => {
    const f = applied;
    if (f.q!=='all'&&r.q!==f.q) return false;
    if (f.t!=='all'&&r.t!==f.t) return false;
    if (f.c!=='all'&&r.c!==f.c) return false;
    if (f.search&&!r.d.toLowerCase().includes(f.search.toLowerCase())) return false;

    // Recorte por intervalo de meses + status (a pagar / liquidado).
    const lo = Math.min(f.monthFrom, f.monthTo);
    const hi = Math.max(f.monthFrom, f.monthTo);
    let sumWindow = 0, sumPaid = 0, sumPending = 0;
    for (let i=lo;i<=hi;i++) {
      const v = r.m[i] || 0;
      sumWindow += v;
      if (paidMonths[i]) sumPaid += v; else sumPending += v;
    }
    if (sumWindow <= 0) return false;
    if (f.status==='liquidado' && sumPaid <= 0) return false;
    if (f.status==='pendente' && sumPending <= 0) return false;
    return true;
  }), [rows, applied, paidMonths]);

  const clearAll = () => { setDraft(initial); setApplied(initial); };

  const startEdit = r => { setEditId(r.id); setEditForm({...r,m:[...r.m]}); };
  const saveEdit = () => {
    const updated = {...editForm, total:editForm.m.reduce((s,v)=>s+v,0)};
    onChange(rows.map(r=>r.id===editId?updated:r));
    setEditId(null); setEditForm(null);
  };
  const delRow = id => {
    if (window.confirm('Excluir esta despesa?')) onChange(rows.filter(r=>r.id!==id));
  };
  const cats = [...new Set(rows.map(r=>r.c))];

  return (
    <div>
      <div style={{ ...s.card, marginBottom:16 }}>
        <p style={{ ...s.label, marginBottom:12 }}>🔎 Filtros avançados</p>
        <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(150px,1fr))', gap:10 }}>
          <Field label="Buscar"><input placeholder="descrição..." value={draft.search} onChange={e=>setDraft({...draft,search:e.target.value})} style={s.input} /></Field>
          <Field label="Pessoa">
            <select value={draft.q} onChange={e=>setDraft({...draft,q:e.target.value})} style={{ ...s.input, cursor:'pointer' }}>
              <option value="all">Todos</option>
              {PESSOAS.map(p=><option key={p} value={p}>{p}</option>)}
            </select>
          </Field>
          <Field label="Forma Pagamento">
            <select value={draft.t} onChange={e=>setDraft({...draft,t:e.target.value})} style={{ ...s.input, cursor:'pointer' }}>
              <option value="all">Todas</option>
              {TIPOS.map(t=><option key={t} value={t}>{t}</option>)}
            </select>
          </Field>
          <Field label="Categoria">
            <select value={draft.c} onChange={e=>setDraft({...draft,c:e.target.value})} style={{ ...s.input, cursor:'pointer' }}>
              <option value="all">Todas</option>
              {cats.map(c=><option key={c} value={c}>{c}</option>)}
            </select>
          </Field>
          <Field label="Mês inicial">
            <select value={draft.monthFrom} onChange={e=>setDraft({...draft,monthFrom:parseInt(e.target.value)})} style={{ ...s.input, cursor:'pointer' }}>
              {ML.map((m,i)=><option key={i} value={i}>{m}</option>)}
            </select>
          </Field>
          <Field label="Mês final">
            <select value={draft.monthTo} onChange={e=>setDraft({...draft,monthTo:parseInt(e.target.value)})} style={{ ...s.input, cursor:'pointer' }}>
              {ML.map((m,i)=><option key={i} value={i}>{m}</option>)}
            </select>
          </Field>
          <Field label="Status">
            <select value={draft.status} onChange={e=>setDraft({...draft,status:e.target.value})} style={{ ...s.input, cursor:'pointer' }}>
              <option value="all">Todos</option>
              <option value="pendente">A pagar</option>
              <option value="liquidado">Liquidado</option>
            </select>
          </Field>
        </div>
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginTop:14, gap:8, flexWrap:'wrap' }}>
          <span style={{ fontSize:12, color:dirty?C.gold:C.muted }}>
            {dirty ? '⚠️ Filtros alterados — clique em Aplicar.' : `${filtered.length}/${rows.length} lançamentos`}
          </span>
          <div style={{ display:'flex', gap:8 }}>
            <button onClick={clearAll} style={s.btnGhost}>✕ Limpar</button>
            <button onClick={()=>setApplied(draft)} disabled={!dirty} style={{ ...s.btn(dirty?C.gold:C.dim), color:dirty?C.bg:C.muted, cursor:dirty?'pointer':'not-allowed' }}>
              ✓ Aplicar filtros
            </button>
          </div>
        </div>
      </div>

      <div style={{ ...s.card, padding:0, overflow:'hidden' }}>
        <div style={{ overflowX:'auto' }}>
          <table style={{ width:'100%', borderCollapse:'collapse' }}>
            <thead>
              <tr style={{ background:C.panel }}>
                {['Tipo','Quem','Para','Categoria','Descrição','Total',''].map((h,i) => (
                  <th key={i} style={{ textAlign:h==='Total'?'right':'left', padding:'12px 14px', fontSize:11, letterSpacing:1, textTransform:'uppercase', color:C.muted, fontWeight:600, borderBottom:`1px solid ${C.border}`, whiteSpace:'nowrap' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.sort((a,b)=>b.total-a.total).map(r => (
                <tr key={r.id} style={{ borderBottom:`1px solid ${C.border}`, background:editId===r.id?C.panel:'transparent' }}>
                  <td style={{ padding:'10px 14px' }}><span style={s.tag(TYPE_C[r.t]||C.muted)}>{r.t}</span></td>
                  <td style={{ padding:'10px 14px', fontSize:13, color:r.q==='Guilherme'?C.blue:C.amber }}>{r.q}</td>
                  <td style={{ padding:'10px 14px', fontSize:13, color:r.p===r.q?C.muted:C.green }}>{r.p}</td>
                  <td style={{ padding:'10px 14px' }}><span style={s.tag(CAT_C[r.c]||C.muted)}>{r.c}</span></td>
                  <td style={{ padding:'10px 14px', fontSize:13, color:C.text }}>{r.d}</td>
                  <td style={{ padding:'10px 14px', fontFamily:'monospace', fontSize:13, color:r.total>0?C.text:C.muted, textAlign:'right' }}>{maskFmt(getWindowSum(r, applied), hide)}</td>
                  <td style={{ padding:'10px 14px', textAlign:'right', whiteSpace:'nowrap' }}>
                    <button onClick={()=>startEdit(r)} style={{ ...s.btnGhost, padding:'4px 10px', marginRight:4 }} aria-label={`Editar ${r.d}`}>✏️</button>
                    <button onClick={()=>delRow(r.id)} style={{ ...s.btnGhost, padding:'4px 10px', borderColor:C.red, color:C.red }} aria-label={`Excluir ${r.d}`}>🗑</button>
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr>
                  <td colSpan={7} style={{ textAlign:'center', padding:30, color:C.muted, fontSize:13 }}>
                    Nenhum lançamento corresponde aos filtros aplicados.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {editId && editForm && (
        <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.7)', display:'flex', alignItems:'center', justifyContent:'center', zIndex:50, padding:20 }}>
          <div style={{ ...s.card, maxWidth:700, width:'100%', maxHeight:'90vh', overflowY:'auto' }}>
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:16 }}>
              <p style={{ ...s.label, color:C.gold }}>✏️ Editar Despesa</p>
              <button onClick={()=>{setEditId(null);setEditForm(null);}} style={s.btnGhost}>✕ Cancelar</button>
            </div>
            <div style={{ display:'grid', gridTemplateColumns:'repeat(2,1fr)', gap:12 }}>
              <Field label="Tipo"><select style={s.input} value={editForm.t} onChange={e=>setEditForm({...editForm,t:e.target.value})}>{TIPOS.map(t=><option key={t} value={t}>{t}</option>)}</select></Field>
              <Field label="Categoria"><select style={s.input} value={editForm.c} onChange={e=>setEditForm({...editForm,c:e.target.value})}>{[...new Set([...CATEGORIAS,editForm.c])].map(c=><option key={c} value={c}>{c}</option>)}</select></Field>
              <Field label="Quem"><select style={s.input} value={editForm.q} onChange={e=>setEditForm({...editForm,q:e.target.value})}>{PESSOAS.map(p=><option key={p} value={p}>{p}</option>)}</select></Field>
              <Field label="Para"><select style={s.input} value={editForm.p} onChange={e=>setEditForm({...editForm,p:e.target.value})}>{PESSOAS.map(p=><option key={p} value={p}>{p}</option>)}</select></Field>
              <Field label="Descrição" span={2}><input style={s.input} value={editForm.d} onChange={e=>setEditForm({...editForm,d:e.target.value})} /></Field>
            </div>
            <p style={{ ...s.label, marginTop:20, marginBottom:8 }}>Valores Mensais</p>
            <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:8 }}>
              {ML.map((mes,i) => (
                <div key={i}>
                  <label style={{ fontSize:10, color:C.muted, marginBottom:2, display:'block' }}>{mes}</label>
                  <input type="text" value={editForm.m[i]} onChange={e=>{const v=parseFloat(String(e.target.value).replace(',','.'))||0;const nm=[...editForm.m];nm[i]=v;setEditForm({...editForm,m:nm});}} style={{ ...s.input, padding:'6px 10px', fontSize:12, fontFamily:'monospace' }} />
                </div>
              ))}
            </div>
            <div style={{ marginTop:16, padding:'10px 14px', background:C.panel, borderRadius:8, fontSize:13, color:C.gold }}>
              ◈ Total: <strong>{fmt(editForm.m.reduce((s,v)=>s+v,0))}</strong>
            </div>
            <div style={{ display:'flex', justifyContent:'flex-end', gap:8, marginTop:16 }}>
              <button onClick={()=>{setEditId(null);setEditForm(null);}} style={s.btnGhost}>Cancelar</button>
              <button onClick={saveEdit} style={s.btn(C.gold)}>💾 Salvar</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── TAB: IA CFO ───
function IATab({ a }) {
  const [insights, setInsights] = useState('');
  const [loading, setLoading] = useState(false);
  const [done, setDone] = useState(false);

  const generate = async () => {
    setLoading(true); setInsights(''); setDone(false);
    const net = a.g2b - a.b2g;
    const prompt = `Você é um CFO sênior especialista em finanças pessoais e split de despesas entre casais. Analise estes dados e retorne insights em português:

TOTAL ANUAL: ${fmt(a.totalGeral)} | ${a.rows.length} lançamentos
Guilherme: ${fmt(a.gTotal)} (${pct(a.gTotal,a.totalGeral)}) | Bruna: ${fmt(a.bTotal)} (${pct(a.bTotal,a.totalGeral)})
G a pagar: ${fmt(a.g2b)} | B a pagar: ${fmt(a.b2g)} | Saldo líquido: ${fmt(Math.abs(net))} ${net>0?'(G→B)':'(B→G)'}
Fixo: ${fmt(a.fixedTotal)} | Variável: ${fmt(a.varTotal)}

CATEGORIAS: ${Object.entries(a.byCat).sort((x,y)=>y[1]-x[1]).map(([k,v])=>`${k}:${fmt(v)}`).join(' | ')}
EVOLUÇÃO: ${a.monthlyActive.map(m=>`${m.mes}:${fmt(m.total)}`).join(' | ')}
TOP 5: ${a.topDesc.slice(0,5).map(([d,v])=>`${d}:${fmt(v)}`).join(' | ')}

Responda com:
## 🔍 DIAGNÓSTICO
## ⚡ ALERTAS (3-4 específicos com valores)
## ⚖️ ANÁLISE DO SPLIT
## 💡 OPORTUNIDADES DE ECONOMIA
## ✅ TOP 5 AÇÕES PRIORITÁRIAS`;

    try {
      const resp = await fetch("https://api.anthropic.com/v1/messages", {
        method:"POST", headers:{"Content-Type":"application/json"},
        body: JSON.stringify({ model:"claude-sonnet-4-20250514", max_tokens:1500, messages:[{role:"user",content:prompt}] })
      });
      const data = await resp.json();
      setInsights((data.content||[]).map(b=>b.text||'').join('') || '❌ Resposta vazia.');
    } catch(e) { setInsights('❌ Erro: '+e.message); }
    setLoading(false); setDone(true);
  };

  const renderMD = text => text.split('\n').map((line,i) => {
    if (line.startsWith('## ')) return <h3 key={i} style={{ color:C.gold, fontSize:14, fontWeight:700, marginTop:24, marginBottom:10, letterSpacing:1, paddingBottom:6, borderBottom:`1px solid ${C.border}` }}>{line.replace('## ','')}</h3>;
    if (line.startsWith('• ')||line.startsWith('- ')) return <p key={i} style={{ color:C.text, fontSize:13, lineHeight:1.7, paddingLeft:12, marginBottom:4 }}>{'▸ '+line.slice(2)}</p>;
    if (/^\d+\./.test(line.trim())) return <p key={i} style={{ color:C.text, fontSize:13, lineHeight:1.7, paddingLeft:12, marginBottom:6, borderLeft:`2px solid ${C.gold}` }}>{line}</p>;
    if (line.trim()) return <p key={i} style={{ color:C.muted, fontSize:13, lineHeight:1.8, marginBottom:8 }}>{line}</p>;
    return <div key={i} style={{ height:4 }} />;
  });

  return (
    <div>
      <div style={{ ...s.card, marginBottom:16, borderColor:C.gold }}>
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', flexWrap:'wrap', gap:12 }}>
          <div>
            <p style={{ ...s.label, color:C.gold }}>✨ CFO AI — Análise Inteligente</p>
            <p style={{ ...s.muted, marginTop:4 }}>Diagnóstico, alertas, split, otimizações e plano de ação</p>
          </div>
          <button onClick={generate} disabled={loading} style={{ ...s.btn(loading?C.dim:C.gold), color:loading?C.muted:C.bg, cursor:loading?'not-allowed':'pointer' }}>
            {loading?'⏳ Analisando...':done?'🔄 Reanalisar':'✨ Gerar Insights'}
          </button>
        </div>
      </div>
      {loading && (
        <div style={{ ...s.card, textAlign:'center', padding:40 }}>
          <div style={{ fontSize:32, marginBottom:12 }}>🧠</div>
          <p style={{ color:C.gold, fontSize:14 }}>Analisando {a.rows.length} lançamentos...</p>
        </div>
      )}
      {insights && !loading && <div style={s.card}>{renderMD(insights)}</div>}
      {!insights && !loading && (
        <div style={{ ...s.card, textAlign:'center', padding:48, borderStyle:'dashed' }}>
          <p style={{ fontSize:48, marginBottom:12 }}>📊</p>
          <p style={{ color:C.text, fontSize:16, marginBottom:8 }}>Análise CFO sob demanda</p>
          <p style={s.muted}>Clique em "Gerar Insights" para análise executiva completa</p>
        </div>
      )}
    </div>
  );
}

// ─── TAB: ATUALIZAR (XLSX) ───
function AtualizarTab({ onData, currentRows }) {
  const [dragging, setDragging] = useState(false);
  const [status, setStatus] = useState('');
  const ref = useRef();

  const parseFile = file => new Promise((res,rej) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result),{type:'array'});
        const ws = wb.Sheets['Dados Brutos']||wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws,{defval:0});
        const result = [];
        for (const row of json) {
          const tipo = row['Tipo']||row['Pago'];
          if (!tipo||tipo==='Tipo'||tipo==='TOTAL GERAL') continue;
          const m = ML.map(k=>parseFloat(row[k])||0);
          result.push({ id:uid(), t:String(tipo), q:String(row['Quem']||''), p:String(row['Pagar Para']||''), c:String(row['Categoria']||'Outros'), d:String(row['Descrição']||''), m, total:m.reduce((s,v)=>s+v,0) });
        }
        res(result.filter(r=>r.q&&r.d));
      } catch(err) { rej(err); }
    };
    reader.onerror = rej;
    reader.readAsArrayBuffer(file);
  });

  const handle = async file => {
    if (!file) return;
    setStatus('⏳ Processando...');
    try {
      const parsed = await parseFile(file);
      if (parsed.length > 0) {
        if (window.confirm(`Encontrei ${parsed.length} lançamentos. Substituir os ${currentRows.length} atuais?`)) {
          onData(parsed); setStatus(`✅ ${parsed.length} lançamentos importados!`);
        } else setStatus('⚠️ Cancelado.');
      } else setStatus('⚠️ Nenhum dado encontrado.');
    } catch(e) { setStatus('❌ Erro: '+e.message); }
  };

  const exportData = () => {
    const wb = XLSX.utils.book_new();
    const data = [['Tipo','Quem','Pagar Para','Categoria','Descrição',...ML,'Total']];
    currentRows.forEach(r => data.push([r.t,r.q,r.p,r.c,r.d,...r.m,r.total]));
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb,ws,'Dados Brutos');
    XLSX.writeFile(wb,`BG_Finance_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  const resetData = () => {
    if (window.confirm(`Restaurar dados originais (${INITIAL.length} lançamentos)?`)) {
      onData(INITIAL); setStatus('🔄 Dados restaurados.');
    }
  };

  return (
    <div>
      <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(260px,1fr))', gap:16, marginBottom:16 }}>
        <div style={s.card}>
          <p style={{ ...s.label, color:C.green, marginBottom:10 }}>📤 Exportar Excel</p>
          <p style={{ ...s.muted, lineHeight:1.6, marginBottom:14 }}>{currentRows.length} lançamentos atuais — inclui edições e adições manuais.</p>
          <button onClick={exportData} style={s.btn(C.green)}>📥 Baixar Excel</button>
        </div>
        <div style={s.card}>
          <p style={{ ...s.label, color:C.red, marginBottom:10 }}>♻️ Restaurar Originais</p>
          <p style={{ ...s.muted, lineHeight:1.6, marginBottom:14 }}>Reverte para os {INITIAL.length} lançamentos iniciais.</p>
          <button onClick={resetData} style={{ ...s.btn(C.red), background:'transparent', color:C.red, border:`1px solid ${C.red}` }}>🔄 Restaurar</button>
        </div>
      </div>
      <div style={{ ...s.card, marginBottom:16 }}>
        <p style={{ ...s.label, marginBottom:6 }}>📂 Importar Excel</p>
        <p style={{ ...s.muted, lineHeight:1.6 }}>Formato: aba "Dados Brutos" com colunas Tipo, Quem, Pagar Para, Categoria, Descrição, Jan-Dez.</p>
      </div>
      <div onDragOver={e=>{e.preventDefault();setDragging(true)}} onDragLeave={()=>setDragging(false)}
        onDrop={e=>{e.preventDefault();setDragging(false);handle(e.dataTransfer.files[0])}}
        onClick={()=>ref.current.click()}
        style={{ border:`2px dashed ${dragging?C.gold:C.border}`, borderRadius:12, padding:60, textAlign:'center', cursor:'pointer', background:dragging?C.goldDim:C.card, transition:'all 0.2s' }}>
        <input ref={ref} type="file" accept=".xlsx,.xls" onChange={e=>handle(e.target.files[0])} style={{ display:'none' }} />
        <p style={{ fontSize:48, marginBottom:16 }}>📂</p>
        <p style={{ color:C.text, fontSize:16, marginBottom:8 }}>Arraste o .xlsx aqui</p>
        <p style={s.muted}>ou clique para selecionar</p>
        {status && <p style={{ marginTop:20, color:status.startsWith('✅')?C.green:status.startsWith('⚠️')?C.gold:status.startsWith('🔄')?C.blue:status.startsWith('⏳')?C.muted:C.red, fontWeight:600 }}>{status}</p>}
      </div>
    </div>
  );
}

// ─── TABS CONFIG ───
const TABS = [
  { id:'overview', label:'Visão Geral', icon:'◈' },
  { id:'encontro', label:'Encontro', icon:'◬' },
    { id:'categorias', label:'Categorias', icon:'◌' },
  { id:'tabela', label:'Tabela', icon:'▦' },
];
const HIDDEN_TABS = [
  { id:'adicionar', label:'Adicionar', icon:'⊕' },
  { id:'ia', label:'IA CFO', icon:'◎' },
  { id:'atualizar', label:'Atualizar', icon:'↻' },
];

const MONTHLY_TABS = ['overview','encontro','categorias'];

// ─── MAIN APP ───
export default function App() {
  const [data, setData] = useState(INITIAL);
  const [tab, setTab] = useState('overview');
  const [loaded, setLoaded] = useState(false);

  // Mês atual dinâmico (timezone do navegador, esperado America/Sao_Paulo).
  const currentMonth = useMemo(() => new Date().getMonth(), []);
  const currentMonthLabel = useMemo(() => getCurrentMonthLabel(), []);

  const [viewMode, setViewMode] = useState('annual');
  const [selectedMonth, setSelectedMonth] = useState(currentMonth);

  const [hidePreviousMonths, setHidePreviousMonths] = useState(false);
  const [showOpsPanel, setShowOpsPanel] = useState(false);

  // Ocultar saldos (eye toggle) — persiste em localStorage.
  const [hideBalances, setHideBalances] = useState(false);
  useEffect(() => { setHideBalances(loadHideBalancesPref()); }, []);
  const toggleHideBalance = useCallback(v => {
    setHideBalances(v);
    saveHideBalancesPref(v);
  }, []);

  // Meses liquidados — persistem em localStorage.
  const [paidMonths, setPaidMonths] = useState({});
  useEffect(() => { setPaidMonths(loadPaidMonths()); }, []);
  const togglePaid = useCallback((idx, paid) => {
    setPaidMonths(prev => {
      const next = { ...prev };
      if (paid) next[idx] = true; else delete next[idx];
      savePaidMonths(next);
      return next;
    });
  }, []);

  const [filters, setFilters] = useState({ person:'all', paymentMethod:'all', category:'all' });

  const a = useMemo(() => recompute(data), [data]);

  // Auto-pula para o primeiro mês ativo na carga inicial se nada estiver selecionado.
  useEffect(() => {
    const firstActive = a.monthly.findIndex(m => m.total > 0);
    if (firstActive >= 0 && a.monthly[selectedMonth]?.total === 0) {
      setSelectedMonth(firstActive);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [loaded]);

  useEffect(() => {
    if (hidePreviousMonths && selectedMonth < currentMonth) {
      setSelectedMonth(currentMonth);
    }
  }, [hidePreviousMonths, selectedMonth, currentMonth]);

  useEffect(() => {
    loadData().then(d => { setData(d); setLoaded(true); });
    const style = document.createElement('style');
    style.textContent = `
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body { background: ${C.bg}; color: ${C.text}; font-family: -apple-system, 'Segoe UI', sans-serif; }
      ::-webkit-scrollbar { width: 4px; height: 4px; }
      ::-webkit-scrollbar-track { background: ${C.panel}; }
      ::-webkit-scrollbar-thumb { background: ${C.border}; border-radius: 3px; }
      ::-webkit-scrollbar-thumb:hover { background: ${C.borderLight}; }
      select option { background: ${C.panel}; color: ${C.text}; }
      input::placeholder { color: ${C.muted}; opacity: 0.7; }
      input:focus, select:focus { border-color: ${C.gold} !important; box-shadow: 0 0 0 2px ${C.goldDim}; }
      button:hover { transform: translateY(-1px); filter: brightness(1.1); }
      button:active { transform: translateY(0) !important; }
      button:focus-visible { outline: 2px solid ${C.gold}; outline-offset: 2px; }
      @media (max-width: 640px) {
        .hide-mobile { display: none !important; }
        .tab-label { display: none; }
      }
    `;
    document.head.appendChild(style);
  }, []);

  const updateData = useCallback(newData => { setData(newData); saveData(newData); }, []);

  const showViewToggle = MONTHLY_TABS.includes(tab);

  // Saldo a liquidar (anual, considerando meses liquidados) para exibição no header.
  const pending = useMemo(() => pendingSettlement(a.monthly, paidMonths), [a.monthly, paidMonths]);

  return (
    <HideCtx.Provider value={hideBalances}>
      <PaidCtx.Provider value={{ paidMonths, togglePaid }}>
        <div style={{ background:C.bg, minHeight:'100vh', color:C.text }}>
          {/* ─── HEADER ─── */}
          <div style={{
            background:C.panel,
            borderBottom:`1px solid ${C.border}`,
            padding:'14px 20px',
            display:'flex',
            justifyContent:'space-between',
            alignItems:'center',
            gap:12,
            flexWrap:'wrap'
          }}>
            <div style={{ display:'flex', alignItems:'center', gap:14, flexWrap:'wrap' }}>
              <div>
                <span style={{ fontSize:18, fontWeight:700, color:C.gold, letterSpacing:2 }}>BG</span>
                <span style={{ fontSize:18, fontWeight:300, color:C.text, letterSpacing:1, marginLeft:6 }}>FINANCE</span>
                <span style={{ fontSize:18, fontWeight:300, color:C.muted, letterSpacing:1, marginLeft:6 }}>MANAGEMENT</span>
                {loaded && <span style={{ fontSize:10, color:C.green, marginLeft:8 }} title="Dados carregados">●</span>}
              </div>
              <div style={{
                display:'inline-flex',
                alignItems:'center',
                gap:6,
                padding:'4px 12px',
                background:C.goldDim,
                border:`1px solid ${C.gold}40`,
                borderRadius:20,
              }}>
                <span style={{ fontSize:12 }}>📅</span>
                <span style={{ fontSize:12, color:C.gold, fontWeight:600 }}>{currentMonthLabel}</span>
              </div>
            </div>
            <div style={{ display:'flex', alignItems:'center', gap:12, flexWrap:'wrap' }}>
              <HideBalancesToggle hide={hideBalances} onChange={toggleHideBalance} />
              <div style={{ textAlign:'right' }}>
                <p style={{ fontSize:10, color:C.muted, letterSpacing:1, textTransform:'uppercase' }}>Saldo a Liquidar</p>
                <p style={{ fontSize:18, fontFamily:'monospace', color:pending.net>0.01?C.red:pending.net<-0.01?C.green:C.gold, fontWeight:700 }}>
                  {maskFmt(Math.abs(pending.net), hideBalances)}
                </p>
              </div>
            </div>
          </div>

          {/* ─── TAB BAR ─── */}
          <div style={{ background:C.panel, borderBottom:`1px solid ${C.border}`, padding:'0 12px', display:'flex', gap:2, overflowX:'auto', scrollbarWidth:'none' }}>
            {TABS.map(t => (
              <button key={t.id} onClick={()=>setTab(t.id)}
                aria-pressed={tab===t.id}
                style={{ padding:'12px 14px', background:'none', border:'none', borderBottom:`2px solid ${tab===t.id?C.gold:'transparent'}`, color:tab===t.id?C.gold:C.muted, cursor:'pointer', fontSize:12, fontWeight:tab===t.id?600:400, transition:'all 0.2s', whiteSpace:'nowrap', display:'flex', alignItems:'center', gap:5 }}>
                <span>{t.icon}</span>
                <span className="tab-label">{t.label}</span>
              </button>
            ))}
            <button onClick={()=>setShowOpsPanel(v=>!v)}
              style={{ marginLeft:'auto', padding:'12px 14px', background:'none', border:'none', color:showOpsPanel?C.gold:C.muted, cursor:'pointer', fontSize:12, whiteSpace:'nowrap' }}>
              ◫ Painel oculto
            </button>
          </div>
          {showOpsPanel && (
            <div style={{ background:C.panel, borderBottom:`1px solid ${C.border}`, padding:'6px 12px', display:'flex', gap:6, overflowX:'auto' }}>
              {HIDDEN_TABS.map(t => (
                <button key={t.id} onClick={()=>setTab(t.id)} style={{ padding:'8px 12px', border:`1px solid ${tab===t.id?C.gold:C.border}`, borderRadius:8, background:tab===t.id?C.goldDim:'transparent', color:tab===t.id?C.gold:C.muted, cursor:'pointer' }}>
                  {t.icon} {t.label}
                </button>
              ))}
            </div>
          )}

          {/* ─── CONTROLS BAR ─── */}
          <div style={{ background:C.panel, borderBottom:`1px solid ${C.border}`, padding:'12px 20px' }}>
            <div style={{ maxWidth:1280, margin:'0 auto', display:'flex', flexDirection:'column', gap:10 }}>
              <div style={{ display:'flex', alignItems:'center', gap:12, flexWrap:'wrap' }}>
                {showViewToggle && (
                  <ViewToggle viewMode={viewMode} onChange={setViewMode} />
                )}
                {showViewToggle && viewMode === 'monthly' && (
                  <HidePreviousToggle
                    hidePreviousMonths={hidePreviousMonths}
                    onChange={setHidePreviousMonths}
                    currentMonth={currentMonth}
                  />
                )}
                <FilterBar filters={filters} onChange={setFilters} rows={data} />
              </div>
              {showViewToggle && viewMode === 'monthly' && (
                <MonthSelector
                  selectedMonth={selectedMonth}
                  onChange={setSelectedMonth}
                  monthly={a.monthly}
                  hidePreviousMonths={hidePreviousMonths}
                  currentMonth={currentMonth}
                  paidMonths={paidMonths}
                />
              )}
            </div>
          </div>

          {/* ─── CONTENT ─── */}
          <div style={{ padding:'20px', maxWidth:1280, margin:'0 auto' }}>
            {tab==='overview' && <OverviewTab a={a} viewMode={viewMode} selectedMonth={selectedMonth} filters={filters} currentMonth={currentMonth} onNavigate={setTab} />}
            {tab==='encontro' && <EncontroTab a={a} viewMode={viewMode} selectedMonth={selectedMonth} currentMonth={currentMonth} />}
            {tab==='categorias' && <CategoriasTab a={a} viewMode={viewMode} selectedMonth={selectedMonth} />}
            {tab==='adicionar' && <AdicionarTab rows={data} onChange={updateData} />}
            {tab==='tabela' && <TabelaTab rows={data} onChange={updateData} />}
            {tab==='ia' && <IATab a={a} />}
            {tab==='atualizar' && <AtualizarTab onData={updateData} currentRows={data} />}
          </div>
        </div>
      </PaidCtx.Provider>
    </HideCtx.Provider>
  );
}
