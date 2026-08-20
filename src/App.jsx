import { useState, useEffect, useMemo } from "react";
import { Search, Building2, Users, Briefcase, ChevronDown, ChevronRight, X, Check, TrendingUp, TrendingDown, LayoutGrid, FilePlus2, BarChart3, ArrowRight, Download } from "lucide-react";
import { signInWithEmailAndPassword, signOut, onAuthStateChanged } from "firebase/auth";
import { auth } from "./firebase.js";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  PieChart,
  Pie,
  Cell,
  LineChart,
  Line,
} from "recharts";

const SEED_DATA = [{"id": 0, "area": "Diretoria", "subarea": "", "funcao": "Gerente Industrial", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 1, "area": "Diretoria", "subarea": "", "funcao": "Gerente  ADM", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 2, "area": "Diretoria", "subarea": "", "funcao": "Gerente  Comercial", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 3, "area": "Diretoria", "subarea": "", "funcao": "Diretor Geral", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 4, "area": "Acabamento", "subarea": "", "funcao": "Coordenador", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 5, "area": "Acabamento", "subarea": "", "funcao": "Supervisor de Acabamento", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 6, "area": "Acabamento", "subarea": "", "funcao": "Operador Líder", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 7, "area": "Acabamento", "subarea": "", "funcao": "Op. de Prensa  I", "autorizadas": 14, "lotadas": 13, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 8, "area": "Acabamento", "subarea": "", "funcao": "Op. de Prensa II", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 9, "area": "Acabamento", "subarea": "", "funcao": "Op. de Prensa III", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 10, "area": "Acabamento", "subarea": "", "funcao": "Op. de papel", "autorizadas": 14, "lotadas": 13, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 11, "area": "Acabamento", "subarea": "", "funcao": "Classificador", "autorizadas": 9, "lotadas": 11, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 12, "area": "Acabamento", "subarea": "", "funcao": "Operador Plastificadora", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 13, "area": "Acabamento", "subarea": "", "funcao": "Aux. de Operação", "autorizadas": 17, "lotadas": 18, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 14, "area": "Acabamento", "subarea": "", "funcao": "Aux. Administrativo", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 15, "area": "Acabamento", "subarea": "", "funcao": "Analista Adm de Produção", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 16, "area": "MDF", "subarea": "", "funcao": "Coordenador", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 17, "area": "MDF", "subarea": "", "funcao": "Supervisor de Produção MDF", "autorizadas": 4, "lotadas": 4, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 18, "area": "MDF", "subarea": "MDF - Prensagem", "funcao": "Operador Líder", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 19, "area": "MDF", "subarea": "MDF - Prensagem", "funcao": "Operador de Sala de Controle III", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 20, "area": "MDF", "subarea": "MDF - Prensagem", "funcao": "Operador de Sala de Controle II", "autorizadas": 2, "lotadas": 2, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 21, "area": "MDF", "subarea": "MDF - Prensagem", "funcao": "Operador de Sala de Controle I", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 22, "area": "MDF", "subarea": "MDF - Prensagem", "funcao": "Operador de Serra I", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 23, "area": "MDF", "subarea": "MDF - Prensagem", "funcao": "Auxiliar de Operação", "autorizadas": 4, "lotadas": 4, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 24, "area": "MDF", "subarea": "MDF - Preparação de Fibras", "funcao": "Operador de Sala de Controle III", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 25, "area": "MDF", "subarea": "MDF - Preparação de Fibras", "funcao": "Operador de Sala de Controle II", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 26, "area": "MDF", "subarea": "MDF - Preparação de Fibras", "funcao": "Aux de Operação", "autorizadas": 7, "lotadas": 6, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 27, "area": "MDF", "subarea": "MDF - Preparação de Cavaco", "funcao": "Operador de Sala de Controle II - Preparação de Cavaco", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 28, "area": "MDF", "subarea": "MDF - Preparação de Cavaco", "funcao": "Operador de Sala de Controle I - Preparação de Cavaco", "autorizadas": 2, "lotadas": 2, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 29, "area": "MDF", "subarea": "MDF - Preparação de Cavaco", "funcao": "Auxiliar de Operação", "autorizadas": 4, "lotadas": 2, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 30, "area": "MDF", "subarea": "MDF - Preparação de Cavaco", "funcao": "Afiador", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 31, "area": "MDF", "subarea": "MDF - Preparação de Cavaco", "funcao": "Operador Líder", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 32, "area": "MDF", "subarea": "MDF - Preparação de Cavaco", "funcao": "Operador de Lixadeira II", "autorizadas": 2, "lotadas": 0, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 33, "area": "MDF", "subarea": "MDF - Preparação de Cavaco", "funcao": "Operador de Lixadeira I", "autorizadas": 5, "lotadas": 7, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 34, "area": "MDF", "subarea": "MDF - Preparação de Cavaco", "funcao": "Classificador", "autorizadas": 3, "lotadas": 2, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 35, "area": "MDF", "subarea": "MDF - Preparação de Cavaco", "funcao": "Auxiliar de Operação", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 36, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "", "funcao": "Coordenador de Segurança, Meio Ambiente, Processos e Qualidade", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 37, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Segurança", "funcao": "Supervisor de Segurança do Trabalho", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 38, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Segurança", "funcao": "Analista de Segurança do Trabalho", "autorizadas": 0, "lotadas": 0, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 39, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Segurança", "funcao": "Assistente de Saúde e Segurança", "autorizadas": 1, "lotadas": 0, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 40, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Segurança", "funcao": "Bombeiro Industrial", "autorizadas": 4, "lotadas": 4, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 41, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Segurança", "funcao": "Técnico de Segurança do Trabalho", "autorizadas": 2, "lotadas": 2, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 42, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Meio Ambiente", "funcao": "Supervisor de Meio Ambiente", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 43, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Meio Ambiente", "funcao": "Assistente de Meio Ambiente", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 44, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Meio Ambiente", "funcao": "Op. de ETE", "autorizadas": 5, "lotadas": 5, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 45, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Meio Ambiente", "funcao": "Auxiliar de Coleta de Resíduos", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 46, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Qualidade", "funcao": "Supervisor de Qualidade", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 47, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Qualidade", "funcao": "Analista de Qualidade", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 48, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Qualidade", "funcao": "Auditor de Processos e Qualidade", "autorizadas": 5, "lotadas": 5, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 49, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Qualidade", "funcao": "Auditor de Ensaios Físicos", "autorizadas": 4, "lotadas": 4, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 50, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Qualidade", "funcao": "Auditor de Insumos II", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 51, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Qualidade", "funcao": "Auditor de Insumos I", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 52, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Processos", "funcao": "Analista de Processos III", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 53, "area": "Segurança, Meio Ambiente, Processos e Qualidade", "subarea": "Processos", "funcao": "Analista de Processos II", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 54, "area": "Manutenção", "subarea": "", "funcao": "Coordenador de Manutenção", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 55, "area": "Manutenção", "subarea": "PCM", "funcao": "Supervisora de Planejamento e Controle de Manutenção", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 56, "area": "Manutenção", "subarea": "PCM", "funcao": "Analista de Planejamento", "autorizadas": 1, "lotadas": 0, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 57, "area": "Manutenção", "subarea": "PCM", "funcao": "Planejador de Manutenção III", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 58, "area": "Manutenção", "subarea": "PCM", "funcao": "Planejador de Manutenção II", "autorizadas": 2, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 59, "area": "Manutenção", "subarea": "PCM", "funcao": "Planejador de Manutenção I", "autorizadas": 0, "lotadas": 0, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 60, "area": "Manutenção", "subarea": "PCM", "funcao": "Aux de Planejamento", "autorizadas": 0, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 61, "area": "Manutenção", "subarea": "PCM", "funcao": "Motorista", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 62, "area": "Manutenção", "subarea": "Lubrificação", "funcao": "Lubrificador Industrial Trainee", "autorizadas": 1, "lotadas": 0, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 63, "area": "Manutenção", "subarea": "Lubrificação", "funcao": "Lubrificador Industrial I", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 64, "area": "Manutenção", "subarea": "Lubrificação", "funcao": "Mecanico Lubrificador II", "autorizadas": 0, "lotadas": 0, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 65, "area": "Manutenção", "subarea": "Lubrificação", "funcao": "Mecanico Lubrificador III", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 66, "area": "Manutenção", "subarea": "Elétrica", "funcao": "Supervisor Elétrica", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 67, "area": "Manutenção", "subarea": "Elétrica", "funcao": "Analista de Automação", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 68, "area": "Manutenção", "subarea": "Elétrica", "funcao": "Técnico de Automação", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 69, "area": "Manutenção", "subarea": "Elétrica", "funcao": "Eletricista III", "autorizadas": 4, "lotadas": 4, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 70, "area": "Manutenção", "subarea": "Elétrica", "funcao": "Eletricista II", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 71, "area": "Manutenção", "subarea": "Elétrica", "funcao": "Eletricista I", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 72, "area": "Manutenção", "subarea": "Elétrica", "funcao": "Eletricista Trainee", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 73, "area": "Manutenção", "subarea": "Mecânica", "funcao": "Supervisor  de Manutenção Mecânica", "autorizadas": 2, "lotadas": 2, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 74, "area": "Manutenção", "subarea": "Mecânica", "funcao": "Mecânico Especializado", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 75, "area": "Manutenção", "subarea": "Mecânica", "funcao": "Mecanico IV", "autorizadas": 4, "lotadas": 4, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 76, "area": "Manutenção", "subarea": "Mecânica", "funcao": "Mecanico III", "autorizadas": 4, "lotadas": 2, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 77, "area": "Manutenção", "subarea": "Mecânica", "funcao": "Mecanico II", "autorizadas": 3, "lotadas": 2, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 78, "area": "Manutenção", "subarea": "Mecânica", "funcao": "Mecânico I", "autorizadas": 3, "lotadas": 4, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 79, "area": "Manutenção", "subarea": "Mecânica", "funcao": "Mecanico Trainee", "autorizadas": 2, "lotadas": 2, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 80, "area": "Manutenção", "subarea": "Fabricação e Montagem", "funcao": "Supervisor de Fabricação e Montagem", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 81, "area": "Manutenção", "subarea": "Fabricação e Montagem", "funcao": "Torneiro Mecanico", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 82, "area": "Manutenção", "subarea": "Fabricação e Montagem", "funcao": "Soldador III", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 83, "area": "Manutenção", "subarea": "Fabricação e Montagem", "funcao": "Mecânico de Manutenção III", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 84, "area": "Manutenção", "subarea": "Fabricação e Montagem", "funcao": "Mecânico de Manutenção II", "autorizadas": 0, "lotadas": 0, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 85, "area": "Manutenção", "subarea": "Fabricação e Montagem", "funcao": "Mecânico de Manutenção I", "autorizadas": 2, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 86, "area": "Manutenção", "subarea": "Fabricação e Montagem", "funcao": "Mecânico Trainee", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 87, "area": "Manutenção", "subarea": "Fabricação e Montagem", "funcao": "Oficial de Manutenção Predial", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 88, "area": "Manutenção", "subarea": "Fabricação e Montagem", "funcao": "Auxiliar de Serviços Gerais", "autorizadas": 2, "lotadas": 2, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 89, "area": "Recursos Humanos / Facilities", "subarea": "", "funcao": "Coordenador", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 90, "area": "Recursos Humanos / Facilities", "subarea": "", "funcao": "Supervisor de DP", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 91, "area": "Recursos Humanos / Facilities", "subarea": "", "funcao": "Analista de RH", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 92, "area": "Recursos Humanos / Facilities", "subarea": "", "funcao": "Assistente de RH", "autorizadas": 1, "lotadas": 0, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 93, "area": "Recursos Humanos / Facilities", "subarea": "", "funcao": "Auxiliar de Seviços Gerais", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 94, "area": "Suprimentos / Almoxarifado / TI", "subarea": "", "funcao": "Coordenador", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 95, "area": "Suprimentos / Almoxarifado / TI", "subarea": "Suprimentos", "funcao": "Analista de Compras", "autorizadas": 1, "lotadas": 0, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 96, "area": "Suprimentos / Almoxarifado / TI", "subarea": "Suprimentos", "funcao": "Analista de Suprimentos", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 97, "area": "Suprimentos / Almoxarifado / TI", "subarea": "Suprimentos", "funcao": "Almoxarife I", "autorizadas": 5, "lotadas": 4, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 98, "area": "Suprimentos / Almoxarifado / TI", "subarea": "TI", "funcao": "Analista de TI", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 99, "area": "Financeiro", "subarea": "", "funcao": "Coordenador Financeiro", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 100, "area": "Financeiro", "subarea": "", "funcao": "Supervisor", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 101, "area": "Financeiro", "subarea": "", "funcao": "Analista Financeiro Senior", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 102, "area": "Financeiro", "subarea": "", "funcao": "Analista Financeiro", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 103, "area": "Financeiro", "subarea": "", "funcao": "Analista de Cobrança", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 104, "area": "Financeiro", "subarea": "", "funcao": "Assistente de Controladoria e Financeiro II", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 105, "area": "Financeiro", "subarea": "", "funcao": "Assistente Financeiro", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 106, "area": "Fiscal / Faturamento / Balança / Controladoria", "subarea": "", "funcao": "Coordenador", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 107, "area": "Fiscal / Faturamento / Balança / Controladoria", "subarea": "Controladoria", "funcao": "Analista de Controladoria", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 108, "area": "Fiscal / Faturamento / Balança / Controladoria", "subarea": "Fiscal", "funcao": "Analista Fiscal II", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 109, "area": "Fiscal / Faturamento / Balança / Controladoria", "subarea": "Fiscal", "funcao": "Assistente Fiscal I", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 110, "area": "Fiscal / Faturamento / Balança / Controladoria", "subarea": "Fiscal", "funcao": "Assistente Fiscal III", "autorizadas": 2, "lotadas": 2, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 111, "area": "Fiscal / Faturamento / Balança / Controladoria", "subarea": "Balança", "funcao": "Auxiliar de Recebimento", "autorizadas": 2, "lotadas": 2, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 112, "area": "Fiscal / Faturamento / Balança / Controladoria", "subarea": "Balança", "funcao": "Assistente Fiscal II", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 113, "area": "Fiscal / Faturamento / Balança / Controladoria", "subarea": "Faturamento", "funcao": "Assistente de Faturamento", "autorizadas": 2, "lotadas": 2, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 114, "area": "Fiscal / Faturamento / Balança / Controladoria", "subarea": "Faturamento", "funcao": "Assistente Fiscal II", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 115, "area": "Pátio de Madeiras / Autos", "subarea": "", "funcao": "Supervisor", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 116, "area": "Pátio de Madeiras / Autos", "subarea": "Pátio de Madeiras", "funcao": "Operador de Máquinas Pesadas III", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 117, "area": "Pátio de Madeiras / Autos", "subarea": "Pátio de Madeiras", "funcao": "Operador de Máquinas Pesadas II", "autorizadas": 8, "lotadas": 8, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 118, "area": "Pátio de Madeiras / Autos", "subarea": "Pátio de Madeiras", "funcao": "Operador de Máquinas Pesadas I", "autorizadas": 7, "lotadas": 7, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 119, "area": "Pátio de Madeiras / Autos", "subarea": "Autos", "funcao": "Mecânico de Autos", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 120, "area": "Pátio de Madeiras / Autos", "subarea": "Autos", "funcao": "Operador Manutentor de Maquinas Pesadas", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 121, "area": "Comercial", "subarea": "", "funcao": "Gerente Comercial", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 122, "area": "Comercial", "subarea": "", "funcao": "Coordenador Comercial", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 123, "area": "Comercial", "subarea": "", "funcao": "Coordenador de Vendas e Marketing", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 124, "area": "Comercial", "subarea": "", "funcao": "Supervisor de Vendas e Marketing", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 125, "area": "Comercial", "subarea": "", "funcao": "Analista de Vendas", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 126, "area": "Comercial", "subarea": "", "funcao": "Assistente de Vendas", "autorizadas": 2, "lotadas": 2, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 127, "area": "Comercial", "subarea": "", "funcao": "Assistente de Vendas II", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 128, "area": "Comercial", "subarea": "", "funcao": "Vendedor", "autorizadas": 4, "lotadas": 4, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 129, "area": "Comercial", "subarea": "", "funcao": "Head of Business Intelligence", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 130, "area": "Expedição/Movimentação Interna", "subarea": "", "funcao": "Supervisor", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 131, "area": "Expedição/Movimentação Interna", "subarea": "Expedição", "funcao": "Líder de Carrregamento", "autorizadas": 3, "lotadas": 3, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 132, "area": "Expedição/Movimentação Interna", "subarea": "Expedição", "funcao": "Analista de PCP", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 133, "area": "Expedição/Movimentação Interna", "subarea": "Expedição", "funcao": "Assistente de PCP", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 134, "area": "Expedição/Movimentação Interna", "subarea": "Expedição", "funcao": "Auxiliar de PCP", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 135, "area": "Expedição/Movimentação Interna", "subarea": "Expedição", "funcao": "Enlonador", "autorizadas": 13, "lotadas": 13, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 136, "area": "Expedição/Movimentação Interna", "subarea": "Movimentação interna de produto", "funcao": "Líder de Movimentação Interna de Produtos", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 137, "area": "Expedição/Movimentação Interna", "subarea": "Movimentação interna de produto", "funcao": "Operador de Empilhadeira III", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 138, "area": "Expedição/Movimentação Interna", "subarea": "Movimentação interna de produto", "funcao": "Operador de Empilhadeira II", "autorizadas": 4, "lotadas": 4, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 139, "area": "Expedição/Movimentação Interna", "subarea": "Movimentação interna de produto", "funcao": "Operador de Empilhadeira I", "autorizadas": 11, "lotadas": 11, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}, {"id": 140, "area": "Administrativo", "subarea": "", "funcao": "Advogado", "autorizadas": 1, "lotadas": 1, "afastados": 0, "terceiro": 0, "motivo": "", "status": "", "ativo": true}];

const STATUS_OPTIONS = ["Não iniciado", "Em divulgação", "Em triagem", "Em entrevistas", "Fechada"];
const STORAGE_KEY = "quadro-lotacao-sudati-v3";
const REQUESTS_KEY = "quadro-lotacao-sudati-solicitacoes-v3";
const TIPO_SOLICITACAO = ["Reposição (saída de colaborador)", "Aumento de quadro (posição nova)"];
const SEED_VAGAS_LOG = [{"id": 0, "mes": "Janeiro", "vaga": "Mecânico de Manutenção III", "setor": "Manutenção Mecânica", "gestor": "Diego Ruher", "motivo": "Substituição Marcelo", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-01-05", "dataFechamento": "2026-01-22", "tempoFechamento": 17, "dataIntegracao": "2026-01-27", "tempoAdmissao": 5, "recrutamento": "RE", "nomeAprovado": "Dionatan (mecânico I)", "observacao": ""}, {"id": 1, "mes": "Janeiro", "vaga": "Enlonador", "setor": "Expedição", "gestor": "Nara Wolff", "motivo": "Substituição Bruno", "entrevistados": 2, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-01-06", "dataFechamento": "2026-01-14", "tempoFechamento": 8, "dataIntegracao": "2026-01-20", "tempoAdmissao": 6, "recrutamento": "RE", "nomeAprovado": "Jaques", "observacao": ""}, {"id": 2, "mes": "Janeiro", "vaga": "Auditor de Ensaios Físicos", "setor": "Qualidade", "gestor": "Cintia Alves", "motivo": "Substituição Nathany", "entrevistados": 1, "mulheres": 1, "situacao": "Fechada", "dataAbertura": "2026-01-19", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "RI", "nomeAprovado": "Yasmim", "observacao": ""}, {"id": 3, "mes": "Janeiro", "vaga": "Auditor de Ensaios Físicos", "setor": "Qualidade", "gestor": "Cintia Alves", "motivo": "Substituição Ana Beatriz", "entrevistados": 4, "mulheres": 3, "situacao": "Fechada", "dataAbertura": "2026-01-19", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "RI/RE", "nomeAprovado": "Emanuel", "observacao": ""}, {"id": 4, "mes": "Janeiro", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Edson Pereira", "motivo": "Substituição Sirlei", "entrevistados": 6, "mulheres": 3, "situacao": "Fechada", "dataAbertura": "2026-01-21", "dataFechamento": "2026-01-22", "tempoFechamento": 1, "dataIntegracao": "2026-01-27", "tempoAdmissao": 5, "recrutamento": "RE", "nomeAprovado": "Evelim", "observacao": ""}, {"id": 5, "mes": "Janeiro", "vaga": "Operador de Papel", "setor": "Acabamento", "gestor": "Edson Pereira", "motivo": "Substituição Gisele", "entrevistados": 6, "mulheres": 3, "situacao": "Fechada", "dataAbertura": "2026-01-21", "dataFechamento": "2026-01-22", "tempoFechamento": 1, "dataIntegracao": "2026-01-27", "tempoAdmissao": 5, "recrutamento": "RE", "nomeAprovado": "Dienefer", "observacao": ""}, {"id": 6, "mes": "Janeiro", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Edson Pereira", "motivo": "Substituição Nicolas", "entrevistados": 5, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-01-21", "dataFechamento": "2026-01-23", "tempoFechamento": 2, "dataIntegracao": "2026-01-27", "tempoAdmissao": 4, "recrutamento": "RE", "nomeAprovado": "Jean", "observacao": ""}, {"id": 7, "mes": "Janeiro", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Rodrigo Pereira", "motivo": "Substituição Elen", "entrevistados": 5, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-01-21", "dataFechamento": "2026-01-23", "tempoFechamento": 2, "dataIntegracao": "2026-01-27", "tempoAdmissao": 4, "recrutamento": "RE", "nomeAprovado": "Gabriel", "observacao": ""}, {"id": 8, "mes": "Janeiro", "vaga": "Enlonador", "setor": "Expedição", "gestor": "Nara Wolff", "motivo": "Substituição José Mateus", "entrevistados": 3, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-01-22", "dataFechamento": "2026-01-29", "tempoFechamento": 7, "dataIntegracao": "2026-02-03", "tempoAdmissao": 5, "recrutamento": "RE", "nomeAprovado": "Vinicius", "observacao": ""}, {"id": 9, "mes": "Janeiro", "vaga": "Aux. Administrativo", "setor": "Acabamento", "gestor": "Paulo André", "motivo": "Substituição Yasmim", "entrevistados": 2, "mulheres": 2, "situacao": "Fechada", "dataAbertura": "2026-01-13", "dataFechamento": "2026-01-14", "tempoFechamento": 1, "dataIntegracao": "2026-01-20", "tempoAdmissao": 6, "recrutamento": "RE", "nomeAprovado": "Mariana", "observacao": ""}, {"id": 10, "mes": "Janeiro", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Leticia Adriano", "motivo": "Substituição Gabriel Kuster", "entrevistados": 6, "mulheres": 3, "situacao": "Fechada", "dataAbertura": "2026-01-21", "dataFechamento": "2026-01-22", "tempoFechamento": 1, "dataIntegracao": "2026-01-27", "tempoAdmissao": 5, "recrutamento": "RE", "nomeAprovado": "Davi", "observacao": ""}, {"id": 11, "mes": "Janeiro", "vaga": "Mecânico Trainee", "setor": "Manutenção Mecânica", "gestor": "Adenilson Mariano", "motivo": "Substituição Anderson", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-01-21", "dataFechamento": "2026-01-22", "tempoFechamento": 1, "dataIntegracao": "2026-01-27", "tempoAdmissao": 5, "recrutamento": "RE", "nomeAprovado": "Daniel", "observacao": ""}, {"id": 12, "mes": "Janeiro", "vaga": "Eletricista Trainee", "setor": "Manutenção Elétrica", "gestor": "Adenilson Mariano", "motivo": "Substituição Maria", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-01-21", "dataFechamento": "2026-01-22", "tempoFechamento": 1, "dataIntegracao": "2026-01-27", "tempoAdmissao": 5, "recrutamento": "RE", "nomeAprovado": "Luciano", "observacao": ""}, {"id": 13, "mes": "Janeiro", "vaga": "Auxiliar de Recebimento", "setor": "Balança", "gestor": "Evandro Quadros", "motivo": "Substituição Bianca", "entrevistados": 2, "mulheres": 1, "situacao": "Fechada", "dataAbertura": "2026-01-16", "dataFechamento": "2026-01-28", "tempoFechamento": 12, "dataIntegracao": "2026-02-03", "tempoAdmissao": 6, "recrutamento": "RE", "nomeAprovado": "Mateus", "observacao": ""}, {"id": 14, "mes": "Janeiro", "vaga": "Estagiário", "setor": "Processos", "gestor": "Lucas Rafael", "motivo": "Substituição Joelma", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-01-09", "dataFechamento": "2026-02-13", "tempoFechamento": 35, "dataIntegracao": "2026-03-03", "tempoAdmissao": 18, "recrutamento": "E", "nomeAprovado": "Mateus Faria de Sousa", "observacao": ""}, {"id": 15, "mes": "Janeiro", "vaga": "Estagiário", "setor": "Processos", "gestor": "Lucas Rafael", "motivo": "Substituição Rafa", "entrevistados": 1, "mulheres": 1, "situacao": "Fechada", "dataAbertura": "2026-01-09", "dataFechamento": "2026-02-11", "tempoFechamento": 33, "dataIntegracao": "2026-02-24", "tempoAdmissao": 13, "recrutamento": "E", "nomeAprovado": "Giovana Daboie", "observacao": ""}, {"id": 16, "mes": "Janeiro", "vaga": "Operador de Empilhadeira", "setor": "Movimentação Interna", "gestor": "Nara Wolff", "motivo": "Substituição Carol", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-01-22", "dataFechamento": "2026-02-04", "tempoFechamento": 13, "dataIntegracao": "2026-02-12", "tempoAdmissao": 8, "recrutamento": "E", "nomeAprovado": "Cleiton Nicoletti", "observacao": ""}, {"id": 17, "mes": "Janeiro", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Maicom William", "motivo": "Substituição Caio", "entrevistados": 5, "mulheres": 3, "situacao": "Fechada", "dataAbertura": "2026-01-27", "dataFechamento": "2026-02-11", "tempoFechamento": 15, "dataIntegracao": "2026-02-24", "tempoAdmissao": 13, "recrutamento": "E", "nomeAprovado": "Jefferson Evangelista da Silva", "observacao": ""}, {"id": 18, "mes": "Janeiro", "vaga": "Auxiliar de Operação", "setor": "Preparação de Fibras", "gestor": "Valdemiro Junior", "motivo": "Substituição Bruno", "entrevistados": 3, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-01-28", "dataFechamento": "2026-02-10", "tempoFechamento": 13, "dataIntegracao": "2026-02-12", "tempoAdmissao": 2, "recrutamento": "E", "nomeAprovado": "Quelven Henrique Fernandes Ortiz da Luz", "observacao": ""}, {"id": 19, "mes": "Fevereiro", "vaga": "Auditor de Ensaios Físicos", "setor": "Qualidade", "gestor": "Cintia Alves", "motivo": "Substituição Marcela", "entrevistados": 3, "mulheres": 1, "situacao": "Fechada", "dataAbertura": "2026-02-02", "dataFechamento": "2026-02-13", "tempoFechamento": 11, "dataIntegracao": "2026-02-19", "tempoAdmissao": 6, "recrutamento": "E", "nomeAprovado": "Ana Carolina Velho", "observacao": ""}, {"id": 20, "mes": "Fevereiro", "vaga": "Auxiliar de Operação", "setor": "Pátio de Madeiras", "gestor": "Everton Pereira Alves", "motivo": "Substituição Adrian", "entrevistados": 3, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-02-03", "dataFechamento": "2026-02-10", "tempoFechamento": 7, "dataIntegracao": "2026-02-12", "tempoAdmissao": 2, "recrutamento": "E", "nomeAprovado": "Jose Luciano Mariano", "observacao": ""}, {"id": 21, "mes": "Fevereiro", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Maicom William", "motivo": "Substituição", "entrevistados": 2, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-02-12", "dataFechamento": "2026-02-13", "tempoFechamento": 1, "dataIntegracao": "2026-02-19", "tempoAdmissao": 6, "recrutamento": "E", "nomeAprovado": "Christian Silva", "observacao": ""}, {"id": 22, "mes": "Fevereiro", "vaga": "Operador de Empilhadeira", "setor": "Movimentação Interna", "gestor": "Nara Wolff", "motivo": "Substituição Mateus Paes", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-02-12", "dataFechamento": "2026-02-13", "tempoFechamento": 1, "dataIntegracao": "2026-02-19", "tempoAdmissao": 6, "recrutamento": "E", "nomeAprovado": "Fabiano Balbinotti", "observacao": ""}, {"id": 23, "mes": "Fevereiro", "vaga": "Auxiliar de Operação", "setor": "Lixadeira", "gestor": "Valdemiro Junior", "motivo": "Substituição Janaina", "entrevistados": 3, "mulheres": 1, "situacao": "Fechada", "dataAbertura": "2026-02-20", "dataFechamento": "2026-02-25", "tempoFechamento": 5, "dataIntegracao": "2026-03-03", "tempoAdmissao": 6, "recrutamento": "E", "nomeAprovado": "Angélica", "observacao": ""}, {"id": 24, "mes": "Março", "vaga": "Auxiliar de Operação", "setor": "Preparação de Fibras", "gestor": "Fernando Moser", "motivo": "Substituição Adair", "entrevistados": 6, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-30", "dataFechamento": "2026-03-31", "tempoFechamento": 1, "dataIntegracao": "2026-04-02", "tempoAdmissao": 2, "recrutamento": "RE", "nomeAprovado": "Ricardo", "observacao": ""}, {"id": 25, "mes": "Março", "vaga": "Auxiliar de Operação", "setor": "Prensagem", "gestor": "Valdemiro Junior", "motivo": "Substituição Eriki", "entrevistados": 6, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-19", "dataFechamento": "2026-03-30", "tempoFechamento": 11, "dataIntegracao": "2026-04-09", "tempoAdmissao": 10, "recrutamento": "RE", "nomeAprovado": "Gustavo", "observacao": ""}, {"id": 26, "mes": "Março", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Rodrigo Pereira", "motivo": "Substituição Evelim", "entrevistados": 3, "mulheres": 3, "situacao": "Fechada", "dataAbertura": "2026-03-26", "dataFechamento": "2026-03-31", "tempoFechamento": 5, "dataIntegracao": "2026-04-02", "tempoAdmissao": 2, "recrutamento": "RE", "nomeAprovado": "Carolina", "observacao": ""}, {"id": 27, "mes": "Março", "vaga": "Operador de Papel", "setor": "Acabamento", "gestor": "Rodrigo Pereira", "motivo": "Substituição Alexandra", "entrevistados": 3, "mulheres": 3, "situacao": "Fechada", "dataAbertura": "2026-03-26", "dataFechamento": "2026-03-31", "tempoFechamento": 5, "dataIntegracao": "2026-04-02", "tempoAdmissao": 2, "recrutamento": "RE", "nomeAprovado": "Ivonete", "observacao": ""}, {"id": 28, "mes": "Março", "vaga": "Auxiliar de Operação", "setor": "Prensagem", "gestor": "Fernando Moser", "motivo": "Substituição Antoni", "entrevistados": 6, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-26", "dataFechamento": "2026-03-31", "tempoFechamento": 5, "dataIntegracao": "2026-04-02", "tempoAdmissao": 2, "recrutamento": "RE", "nomeAprovado": "Diogo", "observacao": ""}, {"id": 29, "mes": "Março", "vaga": "Operador de Empilhadeira", "setor": "Expedição", "gestor": "Nara Wolff", "motivo": "Susbtituição Fabiano", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-27", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "Re", "nomeAprovado": "Adair", "observacao": ""}, {"id": 30, "mes": "Março", "vaga": "Operador de Empilhadeira", "setor": "Expedição", "gestor": "Nara Wolff", "motivo": "Substituição Gilmar", "entrevistados": 2, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-27", "dataFechamento": "2026-04-01", "tempoFechamento": 5, "dataIntegracao": "2026-04-02", "tempoAdmissao": 1, "recrutamento": "RE", "nomeAprovado": "Vandeir", "observacao": ""}, {"id": 31, "mes": "Março", "vaga": "Auxiliar de Operação", "setor": "Lixadeira", "gestor": "Fernando Moser", "motivo": "Substituição Eduardo", "entrevistados": 7, "mulheres": 1, "situacao": "Fechada", "dataAbertura": "2026-03-26", "dataFechamento": "2026-03-31", "tempoFechamento": 5, "dataIntegracao": "2026-04-02", "tempoAdmissao": 2, "recrutamento": "RE", "nomeAprovado": "William", "observacao": ""}, {"id": 32, "mes": "Janeiro", "vaga": "Eletricista II", "setor": "Manutenção Elétrica", "gestor": "Carlos Wolff", "motivo": "Substituição Roberto", "entrevistados": 6, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-01-26", "dataFechamento": "2026-03-18", "tempoFechamento": 51, "dataIntegracao": "2026-04-02", "tempoAdmissao": 15, "recrutamento": "RE", "nomeAprovado": "Igor", "observacao": ""}, {"id": 33, "mes": "Janeiro", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Maicom William", "motivo": "Substituição Caio", "entrevistados": 5, "mulheres": 3, "situacao": "Fechada", "dataAbertura": "2026-01-27", "dataFechamento": "2026-03-10", "tempoFechamento": 42, "dataIntegracao": "2026-03-17", "tempoAdmissao": 7, "recrutamento": "Re", "nomeAprovado": "Marcos", "observacao": ""}, {"id": 34, "mes": "Janeiro", "vaga": "Mecânico III", "setor": "Manutenção Mecânica", "gestor": "Diego Ruher", "motivo": "Substituição Valdemir", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-01-30", "dataFechamento": "2026-03-13", "tempoFechamento": 42, "dataIntegracao": "2026-03-17", "tempoAdmissao": 4, "recrutamento": "RE", "nomeAprovado": "Rogério", "observacao": ""}, {"id": 35, "mes": "Fevereiro", "vaga": "Operador de Empilhadeira", "setor": "Movimentação Interna", "gestor": "Nara Wolff", "motivo": "Substituição Hago", "entrevistados": 2, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-02-16", "dataFechamento": "2026-03-09", "tempoFechamento": 21, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "RI", "nomeAprovado": "Bruno", "observacao": ""}, {"id": 36, "mes": "Fevereiro", "vaga": "Auxiliar de Operação", "setor": "Preparação de Fibras", "gestor": "Everton Pereira Alves", "motivo": "Substituição Igor Cassão", "entrevistados": 3, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-02-24", "dataFechamento": "2026-03-03", "tempoFechamento": 7, "dataIntegracao": "2026-03-05", "tempoAdmissao": 2, "recrutamento": "RE", "nomeAprovado": "Leandro", "observacao": ""}, {"id": 37, "mes": "Março", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Edson Pereira", "motivo": "Substituição Evelim", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-09", "dataFechamento": "2026-03-10", "tempoFechamento": 1, "dataIntegracao": "2026-03-17", "tempoAdmissao": 7, "recrutamento": "RE", "nomeAprovado": "Vinicius", "observacao": ""}, {"id": 38, "mes": "Março", "vaga": "Auxiliar de Operação", "setor": "Preparação de Cavaco", "gestor": "Fernando Moser", "motivo": "Substituição Bruno", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-09", "dataFechamento": "2026-03-12", "tempoFechamento": 3, "dataIntegracao": "2026-03-17", "tempoAdmissao": 5, "recrutamento": "RE", "nomeAprovado": "Carlos", "observacao": ""}, {"id": 39, "mes": "Março", "vaga": "Auxiliar de Operação", "setor": "Preparação de fibras", "gestor": "Valdemiro Junior", "motivo": "Substituição Jackson", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-09", "dataFechamento": "2026-03-10", "tempoFechamento": 1, "dataIntegracao": "2026-03-17", "tempoAdmissao": 7, "recrutamento": "RE", "nomeAprovado": "Ruan", "observacao": ""}, {"id": 40, "mes": "Novembro", "vaga": "Analista de Planejamento", "setor": "PCM", "gestor": "Laura Borges", "motivo": "Substituição Gustavo", "entrevistados": 4, "mulheres": 2, "situacao": "Fechada", "dataAbertura": "2025-11-06", "dataFechamento": "2026-03-06", "tempoFechamento": 120, "dataIntegracao": "2026-03-12", "tempoAdmissao": 6, "recrutamento": "RE", "nomeAprovado": "Vitoria", "observacao": ""}, {"id": 41, "mes": "Março", "vaga": "Supervisor", "setor": "Pátio de Madeiras", "gestor": "Pedro Silva", "motivo": "Substituição Jefferson Barbosa", "entrevistados": 1, "mulheres": 1, "situacao": "Fechada", "dataAbertura": "2026-03-03", "dataFechamento": "2026-03-09", "tempoFechamento": 6, "dataIntegracao": "2026-03-17", "tempoAdmissao": 8, "recrutamento": "RE", "nomeAprovado": "Eder", "observacao": ""}, {"id": 42, "mes": "Março", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Rodrigo Pereira", "motivo": "Substituição Gabriel", "entrevistados": 6, "mulheres": 2, "situacao": "Fechada", "dataAbertura": "2026-03-19", "dataFechamento": "2026-03-20", "tempoFechamento": 1, "dataIntegracao": "2026-03-24", "tempoAdmissao": 4, "recrutamento": "RE", "nomeAprovado": "Mauricio", "observacao": ""}, {"id": 43, "mes": "Março", "vaga": "Lubrificador Trainee", "setor": "Lubrificação", "gestor": "Laura Borges", "motivo": "Substituição Thiago", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-06", "dataFechamento": "2026-03-18", "tempoFechamento": 12, "dataIntegracao": "2026-03-24", "tempoAdmissao": 6, "recrutamento": "RE", "nomeAprovado": "Fabricio Abreu", "observacao": ""}, {"id": 44, "mes": "Março", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Kelvin Poleza", "motivo": "Substituição Jean Carlos", "entrevistados": 6, "mulheres": 2, "situacao": "Fechada", "dataAbertura": "2026-03-18", "dataFechamento": "2026-03-20", "tempoFechamento": 2, "dataIntegracao": "2026-03-24", "tempoAdmissao": 4, "recrutamento": "RE", "nomeAprovado": "Rafaela", "observacao": ""}, {"id": 45, "mes": "Abril", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Maicom William", "motivo": "Substituição Gabriel Chaves", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-23", "dataFechamento": "2026-04-30", "tempoFechamento": 7, "dataIntegracao": "2026-05-06", "tempoAdmissao": 6, "recrutamento": "RE", "nomeAprovado": "Vitor", "observacao": ""}, {"id": 46, "mes": "Abril", "vaga": "Enlonador", "setor": "Expedição", "gestor": "Nara Wolff", "motivo": "Substituição Jonas", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-13", "dataFechamento": "2026-04-27", "tempoFechamento": 14, "dataIntegracao": "2026-04-29", "tempoAdmissao": 2, "recrutamento": "RE", "nomeAprovado": "Uiliam", "observacao": ""}, {"id": 47, "mes": "Março", "vaga": "Mecânico de Manutenção II", "setor": "Manutenção Mecânica", "gestor": "Diego Ruher", "motivo": "Substituição José Fabiano", "entrevistados": 0, "mulheres": 0, "situacao": "Congelada", "dataAbertura": "2026-03-30", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 48, "mes": "Março", "vaga": "Operador de Empilhadeira", "setor": "Movimentação Interna", "gestor": "Nara Wolff", "motivo": "Substituição Gustavo", "entrevistados": 6, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-30", "dataFechamento": "2026-04-15", "tempoFechamento": 16, "dataIntegracao": "2026-04-23", "tempoAdmissao": 8, "recrutamento": "RE", "nomeAprovado": "Gustavo", "observacao": ""}, {"id": 49, "mes": "Abril", "vaga": "Operador de Empilhadeira", "setor": "Movimentação Interna", "gestor": "Nara Wolff", "motivo": "Substituição Isabel", "entrevistados": 3, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-14", "dataFechamento": "2026-04-15", "tempoFechamento": 1, "dataIntegracao": "2026-04-23", "tempoAdmissao": 8, "recrutamento": "RE", "nomeAprovado": "Daniel", "observacao": ""}, {"id": 50, "mes": "Abril", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Rodrigo Pereira", "motivo": "Substituição Cristian", "entrevistados": 6, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-07", "dataFechamento": "2026-04-14", "tempoFechamento": 7, "dataIntegracao": "2026-04-23", "tempoAdmissao": 9, "recrutamento": "RE", "nomeAprovado": "Felipe", "observacao": ""}, {"id": 51, "mes": "Fevereiro", "vaga": "Torneiro Mecânico", "setor": "Fabricação", "gestor": "Emerson Fortuna", "motivo": "Substituição Jean Carlos", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-02-18", "dataFechamento": "2026-04-17", "tempoFechamento": 58, "dataIntegracao": "2026-04-23", "tempoAdmissao": 6, "recrutamento": "RE", "nomeAprovado": "Ismael", "observacao": ""}, {"id": 52, "mes": "Março", "vaga": "Auxiliar de Serviços Gerais", "setor": "Fabricação", "gestor": "Emerson Fortuna", "motivo": "Aumento de quadro", "entrevistados": 2, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-30", "dataFechamento": "2026-04-15", "tempoFechamento": 16, "dataIntegracao": "2026-04-23", "tempoAdmissao": 8, "recrutamento": "RE", "nomeAprovado": "Antonio", "observacao": ""}, {"id": 53, "mes": "Março", "vaga": "Auxiliar de Operação", "setor": "Preparação de Cavaco", "gestor": "Everton Pereira Alves", "motivo": "Substituição William", "entrevistados": 6, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-24", "dataFechamento": "2026-04-07", "tempoFechamento": 14, "dataIntegracao": "2026-04-09", "tempoAdmissao": 2, "recrutamento": "RE", "nomeAprovado": "Iuri", "observacao": ""}, {"id": 54, "mes": "Março", "vaga": "Enlonador", "setor": "Expedição", "gestor": "Nara Wolff", "motivo": "Substituição Guilherme Gerber", "entrevistados": 6, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-27", "dataFechamento": "2026-04-07", "tempoFechamento": 11, "dataIntegracao": "2026-04-09", "tempoAdmissao": 2, "recrutamento": "RE", "nomeAprovado": "Emanuel", "observacao": ""}, {"id": 55, "mes": "Abril", "vaga": "Auxiliar de Operação", "setor": "Prensagem", "gestor": "Everton Pereira Alves", "motivo": "Substituição Lucas", "entrevistados": 6, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-06", "dataFechamento": "2026-04-07", "tempoFechamento": 1, "dataIntegracao": "2026-04-09", "tempoAdmissao": 2, "recrutamento": "RE", "nomeAprovado": "Mayko", "observacao": ""}, {"id": 56, "mes": "Abril", "vaga": "Auxiliar de Operação", "setor": "Preparação de Fibras", "gestor": "Everton Pereira Alves", "motivo": "Substituição Lucas", "entrevistados": 6, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-09", "dataFechamento": "2026-04-13", "tempoFechamento": 4, "dataIntegracao": "2026-04-16", "tempoAdmissao": 3, "recrutamento": "RE", "nomeAprovado": "Antonio", "observacao": ""}, {"id": 57, "mes": "Abril", "vaga": "Mecânico Autos", "setor": "Autos", "gestor": "Éder Sabino", "motivo": "Substituição Adair", "entrevistados": 2, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-04", "dataFechamento": "2026-04-15", "tempoFechamento": 11, "dataIntegracao": "2026-04-23", "tempoAdmissao": 8, "recrutamento": "RE", "nomeAprovado": "André", "observacao": ""}, {"id": 58, "mes": "Abril", "vaga": "Operador de Papel", "setor": "Acabamento", "gestor": "Rodrigo Pereira", "motivo": "Substituição Ivonete", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-28", "dataFechamento": "2026-04-28", "tempoFechamento": 0, "dataIntegracao": "2026-04-29", "tempoAdmissao": 1, "recrutamento": "RE", "nomeAprovado": "Amanda", "observacao": ""}, {"id": 59, "mes": "Março", "vaga": "Eletricista III", "setor": "Manutenção Elétrica", "gestor": "Carlos Wolff", "motivo": "Substituição Gabriel", "entrevistados": 6, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-10", "dataFechamento": "2026-05-20", "tempoFechamento": 71, "dataIntegracao": "2026-05-26", "tempoAdmissao": 6, "recrutamento": "RE", "nomeAprovado": "Rafael", "observacao": ""}, {"id": 60, "mes": "Abril", "vaga": "Eletricista III", "setor": "Manutenção Elétrica", "gestor": "Carlos Wolff", "motivo": "Substituição Luis", "entrevistados": 2, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-22", "dataFechamento": "2026-06-01", "tempoFechamento": 40, "dataIntegracao": "2026-06-09", "tempoAdmissao": 8, "recrutamento": "RE", "nomeAprovado": "Arlen", "observacao": ""}, {"id": 61, "mes": "Abril", "vaga": "Auxiliar de Operação", "setor": "Preparação de Cavaco", "gestor": "Everton Pereira", "motivo": "Substitituição Iure", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-27", "dataFechamento": "2026-05-22", "tempoFechamento": 25, "dataIntegracao": "2026-06-02", "tempoAdmissao": 11, "recrutamento": "RE", "nomeAprovado": "Gladistone", "observacao": ""}, {"id": 62, "mes": "Abril", "vaga": "Auxiliar de Operação", "setor": "Preparação de Fibras", "gestor": "Everton Pereira", "motivo": "Substituição Leandro", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-27", "dataFechamento": "2026-05-30", "tempoFechamento": 33, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "RI", "nomeAprovado": "Leonardo", "observacao": ""}, {"id": 63, "mes": "Maio", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Rodrigo Pereira", "motivo": "Substituição Mauricio", "entrevistados": 3, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-05-13", "dataFechamento": "2026-05-22", "tempoFechamento": 9, "dataIntegracao": "2026-05-26", "tempoAdmissao": 4, "recrutamento": "RE", "nomeAprovado": "José Manoel", "observacao": ""}, {"id": 64, "mes": "Maio", "vaga": "Enlonador", "setor": "Expedição", "gestor": "Nara Wolff", "motivo": "Substituição Jaques", "entrevistados": 3, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-05-13", "dataFechamento": "2026-05-15", "tempoFechamento": 2, "dataIntegracao": "2026-05-19", "tempoAdmissao": 4, "recrutamento": "RE", "nomeAprovado": "Vitor", "observacao": ""}, {"id": 65, "mes": "Março", "vaga": "Auxiliar de Serviços Gerais", "setor": "Fabricação", "gestor": "Emerson Fortuna", "motivo": "Aumento de quadro", "entrevistados": 2, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-03-30", "dataFechamento": "2026-05-05", "tempoFechamento": 36, "dataIntegracao": "2026-05-12", "tempoAdmissao": 7, "recrutamento": "RE", "nomeAprovado": "Cicero", "observacao": ""}, {"id": 66, "mes": "Abril", "vaga": "Auxiliar de Operação", "setor": "Preparação de Fibras", "gestor": "Fernando Moser", "motivo": "Substituição Ricardo", "entrevistados": 8, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-06", "dataFechamento": "2026-05-07", "tempoFechamento": 31, "dataIntegracao": "2026-05-12", "tempoAdmissao": 5, "recrutamento": "RE", "nomeAprovado": "Eltom", "observacao": ""}, {"id": 67, "mes": "Abril", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Maicom William", "motivo": "Substituição Rafael", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-06", "dataFechamento": "2026-05-07", "tempoFechamento": 31, "dataIntegracao": "2026-05-12", "tempoAdmissao": 5, "recrutamento": "RE", "nomeAprovado": "Braian", "observacao": ""}, {"id": 68, "mes": "Abril", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Kelvin Poleza", "motivo": "Substituição Felipe", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-30", "dataFechamento": "2026-05-07", "tempoFechamento": 7, "dataIntegracao": "2026-05-19", "tempoAdmissao": 12, "recrutamento": "RE", "nomeAprovado": "Gabriel", "observacao": ""}, {"id": 69, "mes": "Abril", "vaga": "Auxiliar de Operação", "setor": "Prensagem", "gestor": "Fernando Moser", "motivo": "Substituição Cleison", "entrevistados": 5, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-29", "dataFechamento": "2026-05-04", "tempoFechamento": 5, "dataIntegracao": "2026-05-06", "tempoAdmissao": 2, "recrutamento": "RE", "nomeAprovado": "Vitor", "observacao": ""}, {"id": 70, "mes": "Abril", "vaga": "Auxiliar de Operação", "setor": "Preparação de Cavaco", "gestor": "Valdemiro Junior", "motivo": "Substituição Junior", "entrevistados": 5, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-29", "dataFechamento": "2026-05-04", "tempoFechamento": 5, "dataIntegracao": "2026-05-06", "tempoAdmissao": 2, "recrutamento": "RE", "nomeAprovado": "Leonardo", "observacao": ""}, {"id": 71, "mes": "", "vaga": "Assistente de Vendas", "setor": "Vendas", "gestor": "Jacqueline Akemi", "motivo": "Aumento de quadro", "entrevistados": 1, "mulheres": 1, "situacao": "Fechada", "dataAbertura": "", "dataFechamento": "2026-05-12", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "Erica", "observacao": ""}, {"id": 72, "mes": "Maio", "vaga": "Auxiliar de Operação", "setor": "Lixadeira", "gestor": "Valdemiro Junior", "motivo": "Substituição Angélica", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-05-07", "dataFechamento": "2026-05-15", "tempoFechamento": 8, "dataIntegracao": "2026-05-19", "tempoAdmissao": 4, "recrutamento": "RE", "nomeAprovado": "Tiago", "observacao": ""}, {"id": 73, "mes": "Maio", "vaga": "Operador de Papel", "setor": "Acabamento", "gestor": "Rodrigo Pereira", "motivo": "Substituição Jaine", "entrevistados": 3, "mulheres": 3, "situacao": "Fechada", "dataAbertura": "2026-05-14", "dataFechamento": "2026-05-15", "tempoFechamento": 1, "dataIntegracao": "2026-05-19", "tempoAdmissao": 4, "recrutamento": "RE", "nomeAprovado": "Ariane", "observacao": ""}, {"id": 74, "mes": "Maio", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Kelvin Poleza", "motivo": "Substituição Carolina", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-05-21", "dataFechamento": "2026-05-22", "tempoFechamento": 1, "dataIntegracao": "2026-05-26", "tempoAdmissao": 4, "recrutamento": "RE", "nomeAprovado": "Carlos", "observacao": ""}, {"id": 75, "mes": "Maio", "vaga": "Enlonador", "setor": "Expedição", "gestor": "Nara Wolff", "motivo": "Substituição Uilliam", "entrevistados": 2, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-05-18", "dataFechamento": "2026-05-29", "tempoFechamento": 11, "dataIntegracao": "2026-06-02", "tempoAdmissao": 4, "recrutamento": "RE", "nomeAprovado": "Gabriel", "observacao": ""}, {"id": 76, "mes": "Junho", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Rodrigo Pereira", "motivo": "Substituição José Manoel", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-06-05", "dataFechamento": "2026-06-18", "tempoFechamento": 13, "dataIntegracao": "2026-06-23", "tempoAdmissao": 5, "recrutamento": "RE", "nomeAprovado": "Cleiton", "observacao": ""}, {"id": 77, "mes": "Maio", "vaga": "Operador de Lixadeira", "setor": "Lixadeira", "gestor": "Valdemiro Jr", "motivo": "Substituição Fernando", "entrevistados": 7, "mulheres": 1, "situacao": "Fechada", "dataAbertura": "2026-05-25", "dataFechamento": "2026-06-10", "tempoFechamento": 16, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "RI", "nomeAprovado": "Vinicius Muller", "observacao": ""}, {"id": 78, "mes": "Junho", "vaga": "Classificador", "setor": "Lixadeira", "gestor": "Fernando Moser", "motivo": "Substituição Vinicius", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-06-10", "dataFechamento": "2026-06-10", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "RI", "nomeAprovado": "William Leite", "observacao": ""}, {"id": 79, "mes": "Junho", "vaga": "Auxiliar de Operação", "setor": "Lixadeira", "gestor": "Fernando Moser", "motivo": "Substituição Wiliam Leite", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-06-10", "dataFechamento": "2026-06-10", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "RI", "nomeAprovado": "Ruan Ribeiro", "observacao": ""}, {"id": 80, "mes": "Junho", "vaga": "Enlonador", "setor": "Expedição", "gestor": "Nara Wolff", "motivo": "Substituição Emanuel", "entrevistados": 5, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-06-01", "dataFechamento": "2026-06-03", "tempoFechamento": 2, "dataIntegracao": "2026-06-09", "tempoAdmissao": 6, "recrutamento": "RE", "nomeAprovado": "Edson", "observacao": ""}, {"id": 81, "mes": "Junho", "vaga": "Mecânico de Manutenção I", "setor": "Manutenção Mecânica", "gestor": "Diego Ruher", "motivo": "Substituição Rogério", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-06-02", "dataFechamento": "2026-06-10", "tempoFechamento": 8, "dataIntegracao": "2026-06-16", "tempoAdmissao": 6, "recrutamento": "RE", "nomeAprovado": "Maicon", "observacao": ""}, {"id": 82, "mes": "Junho", "vaga": "Auditor de Ensaios Físicos", "setor": "Qualidade", "gestor": "Cintia Alves", "motivo": "Substituição Daniele P", "entrevistados": 4, "mulheres": 2, "situacao": "Fechada", "dataAbertura": "2026-06-01", "dataFechamento": "2026-06-09", "tempoFechamento": 8, "dataIntegracao": "2026-06-16", "tempoAdmissao": 7, "recrutamento": "RE", "nomeAprovado": "Leandro", "observacao": ""}, {"id": 83, "mes": "Junho", "vaga": "Auditor de Ensaios Físicos", "setor": "Qualidade", "gestor": "Cintia Alves", "motivo": "Substituição Daniele P", "entrevistados": 4, "mulheres": 2, "situacao": "Fechada", "dataAbertura": "2026-06-01", "dataFechamento": "2026-06-09", "tempoFechamento": 8, "dataIntegracao": "2026-06-16", "tempoAdmissao": 7, "recrutamento": "RE", "nomeAprovado": "Vitor", "observacao": ""}, {"id": 84, "mes": "Junho", "vaga": "Auxiliar de Operação", "setor": "Preparação de Cavaco", "gestor": "Valdemiro Jr", "motivo": "Substituição Leonardo", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-06-02", "dataFechamento": "2026-06-18", "tempoFechamento": 16, "dataIntegracao": "2026-06-23", "tempoAdmissao": 5, "recrutamento": "RE", "nomeAprovado": "Roberto", "observacao": ""}, {"id": 85, "mes": "Junho", "vaga": "Mecânico de Manutenção I", "setor": "Manutenção Mecânica", "gestor": "Diego Ruher", "motivo": "Substituição José Fabiano", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-06-02", "dataFechamento": "2026-06-15", "tempoFechamento": 13, "dataIntegracao": "2026-06-23", "tempoAdmissao": 8, "recrutamento": "RE", "nomeAprovado": "Igor José", "observacao": ""}, {"id": 86, "mes": "Junho", "vaga": "Operador de Papel", "setor": "Acabamento", "gestor": "Maicom William", "motivo": "Substituição Tamires", "entrevistados": 4, "mulheres": 4, "situacao": "Fechada", "dataAbertura": "2026-06-18", "dataFechamento": "2026-06-18", "tempoFechamento": 0, "dataIntegracao": "2026-06-23", "tempoAdmissao": 5, "recrutamento": "RE", "nomeAprovado": "Rafaela", "observacao": ""}, {"id": 87, "mes": "Maio", "vaga": "Operador de Sala de Controle", "setor": "Prensagem", "gestor": "Marielle Muniz", "motivo": "Substituição Jonatan", "entrevistados": 7, "mulheres": 1, "situacao": "Fechada", "dataAbertura": "2026-05-25", "dataFechamento": "2026-06-29", "tempoFechamento": 35, "dataIntegracao": "2026-07-07", "tempoAdmissao": 8, "recrutamento": "RE", "nomeAprovado": "Carlos Andrei", "observacao": ""}, {"id": 88, "mes": "Junho", "vaga": "Auxiliar de Operação", "setor": "Prensagem", "gestor": "Fernando Moser", "motivo": "Substituição Diogo", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-06-23", "dataFechamento": "2026-06-29", "tempoFechamento": 6, "dataIntegracao": "2026-07-07", "tempoAdmissao": 8, "recrutamento": "RE", "nomeAprovado": "Cristian", "observacao": ""}, {"id": 89, "mes": "Junho", "vaga": "Auxiliar de Operação", "setor": "Preparação de Fibras", "gestor": "Fernando Moser", "motivo": "Substituição Elton", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-06-26", "dataFechamento": "2026-06-29", "tempoFechamento": 3, "dataIntegracao": "2026-07-07", "tempoAdmissao": 8, "recrutamento": "RE", "nomeAprovado": "Caue", "observacao": ""}, {"id": 90, "mes": "Junho", "vaga": "Auxiliar de Operação", "setor": "Preparação de Fibras", "gestor": "Fernando Moser", "motivo": "Substituição Ruan Ribeiro", "entrevistados": 5, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-06-15", "dataFechamento": "2026-06-29", "tempoFechamento": 14, "dataIntegracao": "2026-07-07", "tempoAdmissao": 8, "recrutamento": "RE", "nomeAprovado": "Igor Medeiros", "observacao": ""}, {"id": 91, "mes": "Junho", "vaga": "Operador de Empilhadeira", "setor": "Movimentação Interna", "gestor": "Nara Wolff", "motivo": "Substituição Cleiton N.", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-06-10", "dataFechamento": "2026-06-23", "tempoFechamento": 13, "dataIntegracao": "2026-07-07", "tempoAdmissao": 14, "recrutamento": "RE", "nomeAprovado": "Matheus", "observacao": ""}, {"id": 92, "mes": "Junho", "vaga": "Enlonador", "setor": "Expedição", "gestor": "Nara Wolff", "motivo": "Substituição Vitor Manoel", "entrevistados": 5, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-06-01", "dataFechamento": "2026-06-29", "tempoFechamento": 28, "dataIntegracao": "2026-07-07", "tempoAdmissao": 8, "recrutamento": "RE", "nomeAprovado": "Guilherme", "observacao": ""}, {"id": 93, "mes": "Abril", "vaga": "Eletricista III", "setor": "Manutenção Elétrica", "gestor": "Carlos Wolff", "motivo": "Substituição Luis", "entrevistados": 3, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-04-22", "dataFechamento": "2026-06-23", "tempoFechamento": 62, "dataIntegracao": "2026-07-07", "tempoAdmissao": 14, "recrutamento": "RE", "nomeAprovado": "Igor José", "observacao": ""}, {"id": 94, "mes": "Novembro", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Maicom William", "motivo": "Substituição Graciele", "entrevistados": 0, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-11-18", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "RI", "nomeAprovado": "Elton", "observacao": ""}, {"id": 95, "mes": "Junho", "vaga": "Eletricista III", "setor": "Manutenção Elétrica", "gestor": "Carlos Wolff", "motivo": "Substituição Irani", "entrevistados": 0, "mulheres": 0, "situacao": "Cancelada", "dataAbertura": "2026-06-29", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 96, "mes": "Junho", "vaga": "Eletricista II", "setor": "Manutenção Elétrica", "gestor": "Carlos Wolff", "motivo": "Substiuição Igor Matheus", "entrevistados": 0, "mulheres": 0, "situacao": "Cancelada", "dataAbertura": "2026-06-30", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 97, "mes": "Julho", "vaga": "Operador de Papel", "setor": "Acabamento", "gestor": "Maicom William", "motivo": "Substituição Rafaela", "entrevistados": 3, "mulheres": 3, "situacao": "Fechada", "dataAbertura": "2026-07-28", "dataFechamento": "2026-07-30", "tempoFechamento": 2, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 98, "mes": "Junho", "vaga": "Auxiliar de Operação", "setor": "Preparação de Cavaco", "gestor": "Valdemiro Jr", "motivo": "Substituição Roberto", "entrevistados": 0, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-06-24", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 99, "mes": "Julho", "vaga": "Op. De Lixadeira", "setor": "Lixadeira", "gestor": "Valdemiro Jr", "motivo": "Substituição Elinton", "entrevistados": 2, "mulheres": 1, "situacao": "Fechada", "dataAbertura": "2026-07-09", "dataFechamento": "2026-07-24", "tempoFechamento": 15, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 100, "mes": "Junho", "vaga": "Auxiliar de Recebimento", "setor": "Balança", "gestor": "Evandro Quadro", "motivo": "Substituição Mateus", "entrevistados": 5, "mulheres": 3, "situacao": "Fechada", "dataAbertura": "2026-06-11", "dataFechamento": "2026-07-10", "tempoFechamento": 29, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 101, "mes": "Julho", "vaga": "Mecanico Especializado", "setor": "Manutenção Mecânica", "gestor": "Diego Ruher", "motivo": "Substituição Alessandro", "entrevistados": 0, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-07-07", "dataFechamento": "2026-07-09", "tempoFechamento": 2, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 102, "mes": "Junho", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Rodrigo Pereira", "motivo": "Mauricio", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-06-24", "dataFechamento": "2026-07-03", "tempoFechamento": 9, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 103, "mes": "Maio", "vaga": "Assistente de Vendas", "setor": "Vendas", "gestor": "Camilla", "motivo": "Substituição Carol", "entrevistados": 1, "mulheres": 1, "situacao": "Fechada", "dataAbertura": "2026-05-07", "dataFechamento": "2026-07-08", "tempoFechamento": 62, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 104, "mes": "Novembro", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Maicom William", "motivo": "Substituição Graciele", "entrevistados": 0, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-11-18", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 105, "mes": "Julho", "vaga": "Auxiliar de Operação", "setor": "Preparação de Fibras", "gestor": "Fernando Moser", "motivo": "Substituição Elton", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-07-01", "dataFechamento": "2026-07-01", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 106, "mes": "Julho", "vaga": "Mecânico Trainee", "setor": "Fabricação", "gestor": "Emerson Fortuna", "motivo": "Substituição Jessica", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-07-10", "dataFechamento": "2026-07-14", "tempoFechamento": 4, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 107, "mes": "Julho", "vaga": "Mecânico Trainee", "setor": "Manutenção Mecânica", "gestor": "Diego Ruher", "motivo": "Substituição Daniel", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-07-10", "dataFechamento": "2026-07-22", "tempoFechamento": 12, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 108, "mes": "Julho", "vaga": "Operador de Papel", "setor": "Acabamento", "gestor": "Rodrigo Pereira", "motivo": "Substituição Mariana", "entrevistados": 2, "mulheres": 2, "situacao": "Fechada", "dataAbertura": "2026-07-10", "dataFechamento": "2026-07-15", "tempoFechamento": 5, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 109, "mes": "Julho", "vaga": "Aux. Expedição - temporário", "setor": "Expedição", "gestor": "Nara Wolf", "motivo": "Substituição Luigi", "entrevistados": 0, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-07-10", "dataFechamento": "2026-07-10", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 110, "mes": "Junho", "vaga": "Mecânico I", "setor": "Fabricação", "gestor": "Emerson Fortuna", "motivo": "Substituição Lauri", "entrevistados": 0, "mulheres": 0, "situacao": "Aberta", "dataAbertura": "2026-06-19", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 111, "mes": "Agosto", "vaga": "Mecânico de Manutenção I", "setor": "Manutenção Mecânica", "gestor": "Diego Ruher", "motivo": "Substituição Dionatan", "entrevistados": 0, "mulheres": 0, "situacao": "Aberta", "dataAbertura": "2026-08-12", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 112, "mes": "Julho", "vaga": "Mecânico de Manutenção III", "setor": "Manutenção Mecânica", "gestor": "Diego Ruher", "motivo": "Substituição Lucas", "entrevistados": 0, "mulheres": 0, "situacao": "Aberta", "dataAbertura": "2026-07-07", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 113, "mes": "Julho", "vaga": "Mecânico Trainee", "setor": "Manutenção Mecânica", "gestor": "Diego Ruher", "motivo": "Substituição Jéssica", "entrevistados": 0, "mulheres": 0, "situacao": "Aberta", "dataAbertura": "2026-07-10", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 114, "mes": "Julho", "vaga": "Técnico de Automação II", "setor": "Manutenção Elétrica", "gestor": "Carlos Wolff", "motivo": "Aumento de quadro", "entrevistados": 0, "mulheres": 0, "situacao": "Aberta", "dataAbertura": "2026-07-29", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 115, "mes": "Agosto", "vaga": "Auxiliar de Operação", "setor": "Acabamento", "gestor": "Maicom William", "motivo": "Substituição Marcos", "entrevistados": 0, "mulheres": 0, "situacao": "Aberta", "dataAbertura": "2026-08-06", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 116, "mes": "Agosto", "vaga": "Auxiliar de Operação", "setor": "Prensagem", "gestor": "Valdemiro Junior", "motivo": "Substituição Vitor", "entrevistados": 0, "mulheres": 0, "situacao": "Aberta", "dataAbertura": "2026-08-06", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 117, "mes": "Agosto", "vaga": "Enlonador", "setor": "Expedição", "gestor": "Nara Wolff", "motivo": "Substituição Gabriel", "entrevistados": 0, "mulheres": 0, "situacao": "Aberta", "dataAbertura": "2026-08-06", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 118, "mes": "Agosto", "vaga": "Enlonador", "setor": "Expedição", "gestor": "Nara Wolff", "motivo": "Aumento de quadro", "entrevistados": 0, "mulheres": 0, "situacao": "Aberta", "dataAbertura": "2026-08-06", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 119, "mes": "Agosto", "vaga": "Operador de Máquinas Pesadas I", "setor": "Pátio de Madeiras", "gestor": "Éder Sabino", "motivo": "Substituição Silvio", "entrevistados": 0, "mulheres": 0, "situacao": "Aberta", "dataAbertura": "2026-08-10", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 120, "mes": "Agosto", "vaga": "Mecânico de máquinas pesadas 1", "setor": "Autos", "gestor": "Éder Sabino", "motivo": "Substituição Iago", "entrevistados": 0, "mulheres": 0, "situacao": "Aberta", "dataAbertura": "2026-08-10", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 121, "mes": "Agosto", "vaga": "Planejador de Manutenção", "setor": "PCM", "gestor": "Laura Farias", "motivo": "Substituição Vitoria", "entrevistados": 0, "mulheres": 0, "situacao": "Aberta", "dataAbertura": "2026-08-10", "dataFechamento": "", "tempoFechamento": 0, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "", "observacao": ""}, {"id": 122, "mes": "Abril", "vaga": "Almoxarife", "setor": "Almoxarifado", "gestor": "Alex Silva", "motivo": "Aumento de quadro", "entrevistados": 11, "mulheres": 4, "situacao": "Fechada", "dataAbertura": "2026-04-14", "dataFechamento": "2026-08-05", "tempoFechamento": 113, "dataIntegracao": "2026-08-11", "tempoAdmissao": 6, "recrutamento": "RE", "nomeAprovado": "Eduardo", "observacao": ""}, {"id": 123, "mes": "Julho", "vaga": "Lubrificador Trainee", "setor": "Lubrificação", "gestor": "Laura Borges", "motivo": "Substituição Fabricio", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-07-16", "dataFechamento": "2026-08-06", "tempoFechamento": 21, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "", "nomeAprovado": "Vitor", "observacao": ""}, {"id": 124, "mes": "Julho", "vaga": "Técnico de Automação II", "setor": "Manutenção Elétrica", "gestor": "Carlos Wolff", "motivo": "Aumento de quadro", "entrevistados": 3, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-07-29", "dataFechamento": "2026-08-11", "tempoFechamento": 13, "dataIntegracao": "2026-09-08", "tempoAdmissao": 28, "recrutamento": "RE", "nomeAprovado": "Leonardo", "observacao": ""}, {"id": 125, "mes": "Julho", "vaga": "Auxiliar de Operação", "setor": "Lixadeira", "gestor": "Fernando Moser", "motivo": "Substituição Vinicius", "entrevistados": 4, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-07-30", "dataFechamento": "2026-08-03", "tempoFechamento": 4, "dataIntegracao": "2026-08-06", "tempoAdmissao": 3, "recrutamento": "RE", "nomeAprovado": "Quelvin", "observacao": ""}, {"id": 126, "mes": "Julho", "vaga": "Operador de Máquinas Pesadas I", "setor": "Pátio de Madeiras", "gestor": "Éder Sabino", "motivo": "Substituição Gabriel", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-07-30", "dataFechamento": "2026-07-30", "tempoFechamento": 0, "dataIntegracao": "2026-08-06", "tempoAdmissao": 7, "recrutamento": "RE", "nomeAprovado": "Silvio", "observacao": ""}, {"id": 127, "mes": "Agosto", "vaga": "Almoxarife", "setor": "Almoxarifado", "gestor": "Alex Silva", "motivo": "Substituição João", "entrevistados": 1, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "2026-08-03", "dataFechamento": "2026-08-05", "tempoFechamento": 2, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "RI", "nomeAprovado": "Marcos", "observacao": ""}, {"id": 128, "mes": "", "vaga": "Técnico de Automação I", "setor": "Manutenção Elétrica", "gestor": "Carlos Wolff", "motivo": "Substituição Rodirgo", "entrevistados": 0, "mulheres": 0, "situacao": "Fechada", "dataAbertura": "", "dataFechamento": "2026-08-11", "tempoFechamento": 46245, "dataIntegracao": "", "tempoAdmissao": 0, "recrutamento": "RE", "nomeAprovado": "Marlon", "observacao": "Vaga sigilosa"}];
const VAGAS_LOG_KEY = "quadro-lotacao-sudati-vagas-log-v2";

function exportBackupToExcel(data, vagasLog, requests) {
  const wb = XLSX.utils.book_new();

  const quadroRows = data.map((r) => ({
    Área: r.area,
    Subárea: r.subarea || "",
    Função: r.funcao,
    Autorizadas: r.autorizadas,
    Lotadas: r.lotadas,
    Afastados: r.afastados,
    Terceiro: r.terceiro,
    "Vagas em aberto": vagasAbertas(r),
    Motivo: r.motivo || "",
    Observação: r.observacaoQuadro || "",
    Ativo: r.ativo === false ? "Não" : "Sim",
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(quadroRows), "Quadro de Lotação");

  const vagasRows = vagasLog.map((v) => ({
    Mês: v.mes,
    Vaga: v.vaga,
    Setor: v.setor,
    Gestor: v.gestor || "",
    Motivo: v.motivo || "",
    Entrevistados: v.entrevistados || 0,
    Mulheres: v.mulheres || 0,
    Situação: v.situacao,
    "Data abertura": v.dataAbertura || "",
    "Data fechamento": v.dataFechamento || "",
    "Tempo fechamento (dias)": v.tempoFechamento || "",
    Integração: v.dataIntegracao || "",
    "Tempo admissão (dias)": v.tempoAdmissao || "",
    Recrutamento: v.recrutamento || "",
    "Nome aprovado": v.nomeAprovado || "",
    Observação: v.observacao || "",
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(vagasRows), "Vagas");

  const solicitacoesRows = requests.map((r) => ({
    Área: r.area,
    Função: r.funcao,
    Tipo: r.tipo,
    Solicitante: r.solicitante || "",
    "Colaborador substituído": r.colaboradorSubstituido || "",
    Justificativa: r.justificativa || "",
    "Data abertura": r.dataAbertura || "",
    Status: r.status,
    "Criada em": r.criadaEm || "",
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(solicitacoesRows), "Solicitações");

  const hoje = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `backup-quadro-lotacao-sudati-${hoje}.xlsx`);
}

function vagasAbertas(row) {
  return Math.max(0, row.autorizadas - row.lotadas - row.terceiro);
}

function ocupacaoPct(rows) {
  const aut = rows.reduce((s, r) => s + r.autorizadas, 0);
  const ocu = rows.reduce((s, r) => s + r.lotadas + r.terceiro, 0);
  return aut === 0 ? 0 : Math.round((ocu / aut) * 100);
}

export default function QuadroLotacao() {
  const [data, setData] = useState(SEED_DATA);
  const [ready, setReady] = useState(false);
  const [tab, setTab] = useState("adminhome");
  const [search, setSearch] = useState("");
  const [areaFilter, setAreaFilter] = useState("Todas");
  const [editingCell, setEditingCell] = useState(null);
  const [expandedAreas, setExpandedAreas] = useState({});
  const [saveState, setSaveState] = useState("idle");
  const [role, setRole] = useState("gestor");
  const [gestorArea, setGestorArea] = useState("");
  const [gestorNome, setGestorNome] = useState("");
  const [requests, setRequests] = useState([]);
  const [vagasLog, setVagasLog] = useState(SEED_VAGAS_LOG);
  const [passwordModal, setPasswordModal] = useState(false);
  const [emailInput, setEmailInput] = useState("");
  const [passwordInput, setPasswordInput] = useState("");
  const [passwordError, setPasswordError] = useState(false);
  const [passwordVisible, setPasswordVisible] = useState(false);
  const [authChecking, setAuthChecking] = useState(true);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (user) => {
      if (user) setRole("admin");
      setAuthChecking(false);
    });
    return unsubscribe;
  }, []);

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, false);
        if (res && res.value) {
          const parsed = JSON.parse(res.value);
          if (Array.isArray(parsed) && parsed.length) setData(parsed);
        } else {
          await window.storage.set(STORAGE_KEY, JSON.stringify(SEED_DATA), false);
        }
      } catch (e) {
        // no stored data yet, keep seed
      }
      try {
        const resReq = await window.storage.get(REQUESTS_KEY, false);
        if (resReq && resReq.value) {
          const parsedReq = JSON.parse(resReq.value);
          if (Array.isArray(parsedReq)) setRequests(parsedReq);
        }
      } catch (e) {
        // no requests yet
      }
      try {
        const resVagas = await window.storage.get(VAGAS_LOG_KEY, false);
        if (resVagas && resVagas.value) {
          const parsedVagas = JSON.parse(resVagas.value);
          if (Array.isArray(parsedVagas) && parsedVagas.length) setVagasLog(parsedVagas);
        } else {
          await window.storage.set(VAGAS_LOG_KEY, JSON.stringify(SEED_VAGAS_LOG), false);
        }
      } catch (e) {
        // keep seed
      }
      setReady(true);
    })();
  }, []);

  const persistVagasLog = async (next) => {
    setVagasLog(next);
    setSaveState("saving");
    try {
      await window.storage.set(VAGAS_LOG_KEY, JSON.stringify(next), false);
      setSaveState("saved");
      setTimeout(() => setSaveState("idle"), 1200);
    } catch (e) {
      setSaveState("error");
    }
  };

  const updateVaga = (id, field, value) => {
    const next = vagasLog.map((r) => (r.id === id ? { ...r, [field]: value } : r));
    persistVagasLog(next);
  };


  const persist = async (next) => {
    setData(next);
    setSaveState("saving");
    try {
      await window.storage.set(STORAGE_KEY, JSON.stringify(next), false);
      setSaveState("saved");
      setTimeout(() => setSaveState("idle"), 1200);
    } catch (e) {
      setSaveState("error");
    }
  };

  const persistRequests = async (next) => {
    setRequests(next);
    setSaveState("saving");
    try {
      await window.storage.set(REQUESTS_KEY, JSON.stringify(next), false);
      setSaveState("saved");
      setTimeout(() => setSaveState("idle"), 1200);
    } catch (e) {
      setSaveState("error");
    }
  };

  const updateRow = (id, field, value) => {
    const next = data.map((r) => (r.id === id ? { ...r, [field]: value } : r));
    persist(next);
  };

  const addPosition = (row) => {
    const newId = Math.max(0, ...data.map((r) => r.id)) + 1;
    const next = [
      ...data,
      {
        id: newId,
        area: row.area,
        subarea: row.subarea || "",
        funcao: row.funcao,
        autorizadas: row.autorizadas,
        lotadas: 0,
        afastados: 0,
        terceiro: 0,
        motivo: "",
        status: "",
        ativo: true,
      },
    ];
    persist(next);
  };

  const removePosition = (id) => {
    const next = data.map((r) => (r.id === id ? { ...r, ativo: false } : r));
    persist(next);
  };

  const restorePosition = (id) => {
    const next = data.map((r) => (r.id === id ? { ...r, ativo: true } : r));
    persist(next);
  };

  const tryUnlockAdmin = async () => {
    setPasswordError(false);
    try {
      await signInWithEmailAndPassword(auth, emailInput.trim(), passwordInput);
      setRole("admin");
      setPasswordModal(false);
      setEmailInput("");
      setPasswordInput("");
    } catch (e) {
      setPasswordError(true);
    }
  };

  const logoutAdmin = async () => {
    await signOut(auth);
    setRole("gestor");
  };

  const submitRequest = (req) => {
    const next = [
      ...requests,
      {
        id: Date.now(),
        criadaEm: new Date().toISOString(),
        status: "Pendente",
        ...req,
      },
    ];
    persistRequests(next);
  };

  const resolveRequest = (id, decision) => {
    const req = requests.find((r) => r.id === id);
    if (!req) return;
    const hoje = new Date().toISOString().slice(0, 10);
    const dataAberturaFinal = req.dataAbertura || hoje;

    if (decision === "Aprovada") {
      const existing = data.find((r) => r.area === req.area && r.funcao === req.funcao);
      if (req.tipo === TIPO_SOLICITACAO[1]) {
        // Aumento de quadro: incrementa autorizadas ou cria nova posição
        if (existing) {
          const next = data.map((r) =>
            r.id === existing.id
              ? { ...r, autorizadas: r.autorizadas + 1, dataAbertura: r.dataAbertura || dataAberturaFinal, gestorResponsavel: req.solicitante || r.gestorResponsavel || "" }
              : r
          );
          persist(next);
        } else {
          const newId = Math.max(0, ...data.map((r) => r.id)) + 1;
          const next = [
            ...data,
            {
              id: newId,
              area: req.area,
              subarea: "",
              funcao: req.funcao,
              autorizadas: 1,
              lotadas: 0,
              afastados: 0,
              terceiro: 0,
              motivo: "",
              status: "Em divulgação",
              ativo: true,
              dataAbertura: dataAberturaFinal,
              dataEncerramento: "",
              dataInicio: "",
              entrevistados: 0,
              entrevistadasMulheres: 0,
              gestorResponsavel: req.solicitante || "",
            },
          ];
          persist(next);
        }
      } else if (existing) {
        // Reposição: já existe como vaga aberta (autorizadas > lotadas); só marca status
        const next = data.map((r) =>
          r.id === existing.id
            ? {
                ...r,
                status: r.status || "Em divulgação",
                dataAbertura: r.dataAbertura || dataAberturaFinal,
                gestorResponsavel: req.solicitante || r.gestorResponsavel || "",
              }
            : r
        );
        persist(next);
      }

      const meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
      ];
      const newVagaId = Math.max(0, ...vagasLog.map((v) => v.id)) + 1;
      const novaVaga = {
        id: newVagaId,
        mes: meses[new Date(dataAberturaFinal).getMonth()] || meses[new Date().getMonth()],
        vaga: req.funcao,
        setor: req.area,
        gestor: req.solicitante || "",
        motivo: req.colaboradorSubstituido ? `Substituição de ${req.colaboradorSubstituido}` : req.justificativa || "",
        entrevistados: 0,
        mulheres: 0,
        situacao: "Aberta",
        dataAbertura: dataAberturaFinal,
        dataFechamento: "",
        tempoFechamento: 0,
        dataIntegracao: "",
        tempoAdmissao: 0,
        recrutamento: "",
        nomeAprovado: "",
        observacao: "",
      };
      persistVagasLog([...vagasLog, novaVaga]);
    }

    const nextRequests = requests.map((r) => (r.id === id ? { ...r, status: decision, resolvidaEm: new Date().toISOString() } : r));
    persistRequests(nextRequests);
  };

  useEffect(() => {
    const adminTabs = ["adminhome", "dashboard", "lotacao", "organograma", "vagas", "indicadores", "solicitacoes"];
    const gestorTabs = ["gestorhome", "meusetor", "vagasgestor", "abrirvaga", "indicadoresgestor"];
    if (role === "admin" && !adminTabs.includes(tab)) setTab("adminhome");
    if (role === "gestor" && !gestorTabs.includes(tab)) setTab("gestorhome");
  }, [role]); // eslint-disable-line react-hooks/exhaustive-deps

  const areas = useMemo(() => {
    const seen = [];
    data.forEach((r) => {
      if (!seen.includes(r.area)) seen.push(r.area);
    });
    return seen;
  }, [data]);

  const filtered = useMemo(() => {
    return data.filter((r) => {
      if (r.ativo === false) return false;
      const matchesArea = areaFilter === "Todas" || r.area === areaFilter;
      const q = search.trim().toLowerCase();
      const matchesSearch =
        !q || r.funcao.toLowerCase().includes(q) || r.area.toLowerCase().includes(q);
      return matchesArea && matchesSearch;
    });
  }, [data, areaFilter, search]);

  const openPositions = useMemo(
    () => filtered.filter((r) => vagasAbertas(r) > 0).sort((a, b) => a.area.localeCompare(b.area)),
    [filtered]
  );

  const totals = useMemo(() => {
    const ativos = data.filter((r) => r.ativo !== false);
    const aut = ativos.reduce((s, r) => s + r.autorizadas, 0);
    const lot = ativos.reduce((s, r) => s + r.lotadas, 0);
    const terc = ativos.reduce((s, r) => s + r.terceiro, 0);
    const afast = ativos.reduce((s, r) => s + r.afastados, 0);
    const abertas = ativos.reduce((s, r) => s + vagasAbertas(r), 0);
    return { aut, lot, terc, afast, abertas, pct: ocupacaoPct(ativos) };
  }, [data]);

  const byArea = useMemo(() => {
    const map = {};
    data
      .filter((r) => r.ativo !== false)
      .forEach((r) => {
        if (!map[r.area]) map[r.area] = [];
        map[r.area].push(r);
      });
    return map;
  }, [data]);

  const toggleArea = (area) => {
    setExpandedAreas((prev) => ({ ...prev, [area]: !prev[area] }));
  };

  const pendingCount = requests.filter((r) => r.status === "Pendente").length;
  const vagasAtivasCount = vagasLog.filter((v) => v.situacao === "Aberta").length;
  const gestorRows = useMemo(() => data.filter((r) => r.area === gestorArea), [data, gestorArea]);
  const gestorOpen = useMemo(() => gestorRows.filter((r) => vagasAbertas(r) > 0), [gestorRows]);
  const gestorRequests = useMemo(() => requests.filter((r) => r.area === gestorArea), [requests, gestorArea]);
  const vagasLogGestor = useMemo(() => {
    if (!gestorArea) return [];
    const alvo = gestorArea.toLowerCase();
    return vagasLog.filter((v) => v.setor && (v.setor.toLowerCase().includes(alvo) || alvo.includes(v.setor.toLowerCase())));
  }, [vagasLog, gestorArea]);

  if (!ready || authChecking) {
    return (
      <div style={{ ...styles.page, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <div style={{ color: C.slate, fontFamily: F.body }}>Carregando quadro de lotação…</div>
      </div>
    );
  }

  return (
    <div style={styles.page}>
      <style>{globalCss}</style>
      <Header
        totals={totals}
        vagasAtivasCount={vagasAtivasCount}
        saveState={saveState}
        role={role}
        setRole={setRole}
        gestorArea={gestorArea}
        setGestorArea={setGestorArea}
        areas={areas}
        passwordModal={passwordModal}
        setPasswordModal={setPasswordModal}
        emailInput={emailInput}
        setEmailInput={setEmailInput}
        passwordInput={passwordInput}
        setPasswordInput={setPasswordInput}
        passwordError={passwordError}
        passwordVisible={passwordVisible}
        setPasswordVisible={setPasswordVisible}
        tryUnlockAdmin={tryUnlockAdmin}
        logoutAdmin={logoutAdmin}
        onExport={() => exportBackupToExcel(data, vagasLog, requests)}
      />
      {!((role === "gestor" && (!gestorArea || tab === "gestorhome")) || (role === "admin" && tab === "adminhome")) && (
        <TabBar tab={tab} setTab={setTab} counts={{ vagas: vagasAtivasCount, areas: areas.length, pendentes: pendingCount }} role={role} />
      )}

      <div style={styles.content}>
        {role === "admin" && tab === "adminhome" && (
          <AdminHome setTab={setTab} vagasCount={vagasAtivasCount} pendentesCount={pendingCount} />
        )}

        {role === "admin" && tab === "dashboard" && <Dashboard totals={totals} byArea={byArea} />}

        {role === "admin" && tab === "lotacao" && (
          <LotacaoTable
            areas={areas}
            areaFilter={areaFilter}
            setAreaFilter={setAreaFilter}
            search={search}
            setSearch={setSearch}
            filtered={filtered}
            allData={data}
            editingCell={editingCell}
            setEditingCell={setEditingCell}
            updateRow={updateRow}
            addPosition={addPosition}
            removePosition={removePosition}
            restorePosition={restorePosition}
          />
        )}

        {role === "admin" && tab === "organograma" && (
          <Organograma
            byArea={byArea}
            areas={areas}
            expandedAreas={expandedAreas}
            toggleArea={toggleArea}
            updateRow={updateRow}
          />
        )}

        {role === "admin" && tab === "vagas" && <VagasAbertas vagasLog={vagasLog} updateVaga={updateVaga} />}

        {role === "admin" && tab === "indicadores" && <Indicadores vagasLog={vagasLog} />}

        {role === "admin" && tab === "solicitacoes" && (
          <Solicitacoes requests={requests} resolveRequest={resolveRequest} />
        )}

        {role === "gestor" && !gestorArea && (
          <div style={{ textAlign: "center", padding: "60px 20px", color: C.slateLight }}>
            Selecione seu setor no topo da página pra ver os dados do seu time.
          </div>
        )}

        {role === "gestor" && gestorArea && tab === "gestorhome" && (
          <GestorHome area={gestorArea} setTab={setTab} vagasCount={gestorOpen.length} />
        )}

        {role === "gestor" && gestorArea && tab === "meusetor" && (
          <MeuSetor area={gestorArea} rows={gestorRows} />
        )}

        {role === "gestor" && gestorArea && tab === "vagasgestor" && (
          <VagasGestor rows={gestorOpen} />
        )}

        {role === "gestor" && gestorArea && tab === "abrirvaga" && (
          <AbrirVaga area={gestorArea} rows={gestorRows} requests={gestorRequests} submitRequest={submitRequest} gestorNome={gestorNome} setGestorNome={setGestorNome} />
        )}

        {role === "gestor" && gestorArea && tab === "indicadoresgestor" && (
          <Indicadores vagasLog={vagasLogGestor} />
        )}
      </div>
    </div>
  );
}

// ---------- design tokens ----------
const C = {
  bg: "#F3F5F6",
  surface: "#FFFFFF",
  ink: "#232A2F",
  wood: "#576978",
  woodDark: "#232A2F",
  slate: "#576978",
  slateLight: "#8B99A3",
  signal: "#4B83AD",
  signalBg: "#E8F1F7",
  moss: "#4B8A6F",
  mossBg: "#E7F3EC",
  line: "#DCE1E4",
  amber: "#B7862E",
  amberBg: "#F7EFDC",
};
const F = {
  display: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif",
  body: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif",
  mono: "'SF Mono', 'Roboto Mono', ui-monospace, monospace",
};

const globalCss = `
  * { box-sizing: border-box; }
  ::selection { background: ${C.wood}33; }
  .qlt-scroll::-webkit-scrollbar { height: 8px; width: 8px; }
  .qlt-scroll::-webkit-scrollbar-thumb { background: ${C.line}; border-radius: 4px; }
  input:focus, select:focus, button:focus-visible { outline: 2px solid ${C.wood}; outline-offset: 1px; }
  button { transition: background 0.15s ease, border-color 0.15s ease, opacity 0.15s ease, transform 0.1s ease; }
  button:active { transform: translateY(1px); }
  .qlt-tab { position: relative; }
  .qlt-tab:hover { color: ${C.ink} !important; }
  .qlt-pill:hover { border-color: ${C.signal} !important; }
  .qlt-row:hover { background: ${C.bg}; }
  .qlt-editable:hover { background: ${C.bg}; }
  .qlt-card-btn:hover { border-color: ${C.signal} !important; box-shadow: 0 4px 14px rgba(22,50,74,0.12) !important; transform: translateY(-2px); }
  @media (max-width: 640px) {
    .qlt-hide-mobile { display: none !important; }
  }
`;

const shadowSm = "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)";

const styles = {
  page: {
    minHeight: "100%",
    background: C.bg,
    fontFamily: F.body,
    color: C.ink,
    display: "flex",
    flexDirection: "column",
  },
  content: {
    padding: "20px 24px 40px",
    maxWidth: 1200,
    margin: "0 auto",
    width: "100%",
  },
};

// ---------- Header ----------
function Header({
  totals,
  vagasAtivasCount,
  saveState,
  role,
  setRole,
  gestorArea,
  setGestorArea,
  areas,
  passwordModal,
  setPasswordModal,
  emailInput,
  setEmailInput,
  passwordInput,
  setPasswordInput,
  passwordError,
  passwordVisible,
  setPasswordVisible,
  tryUnlockAdmin,
  logoutAdmin,
  onExport,
}) {
  return (
    <div
      style={{
        background: C.woodDark,
        color: "#F3F5F6",
        padding: "18px 24px",
        borderBottom: `4px solid ${C.signal}`,
        boxShadow: "0 2px 10px rgba(0,0,0,0.15)",
        position: "relative",
      }}
    >
      <div style={{ maxWidth: 1200, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <BrandBlock />
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap" }}>
          <StatChip label="Ocupação" value={`${totals.pct}%`} />
          <StatChip label="Vagas abertas" value={vagasAtivasCount} accent />
          {role === "admin" && (
            <button
              onClick={onExport}
              title="Baixar todos os dados em Excel"
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                padding: "7px 12px",
                borderRadius: 20,
                border: "1px solid #3E4A52",
                background: "transparent",
                color: "#F3F5F6",
                fontSize: 12.5,
                fontWeight: 600,
                cursor: "pointer",
              }}
            >
              <Download size={14} /> Exportar Excel
            </button>
          )}
          <SaveIndicator state={saveState} />
        </div>
      </div>
      <div style={{ maxWidth: 1200, margin: "10px auto 0", display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
        <RolePill active={role === "admin"} onClick={() => (role === "admin" ? null : setPasswordModal(true))} label="RH (administrador)" />
        <RolePill active={role === "gestor"} onClick={() => setRole("gestor")} label="Gestor de área" />
        {role === "admin" && (
          <button
            onClick={logoutAdmin}
            style={{ padding: "5px 12px", borderRadius: 20, border: "1px solid #3E4A52", background: "transparent", color: "#A8B4BB", fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}
          >
            Sair
          </button>
        )}
        {role === "gestor" && (
          <>
            <select
              value={gestorArea}
              onChange={(e) => setGestorArea(e.target.value)}
              style={{
                padding: "5px 10px",
                borderRadius: 20,
                border: "1px solid #3E4A52",
                background: "#33404A",
                color: "#F3F5F6",
                fontSize: 12.5,
              }}
            >
              <option value="">Selecione seu setor…</option>
              {areas.map((a) => (
                <option key={a} value={a}>
                  {a}
                </option>
              ))}
            </select>
          </>
        )}
      </div>

      {passwordModal && (
        <div
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(20,15,8,0.55)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            zIndex: 50,
          }}
          onClick={() => setPasswordModal(false)}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{ background: C.surface, borderRadius: 8, padding: 22, width: 280, color: C.ink }}
          >
            <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 15, marginBottom: 10 }}>
              Acesso do RH
            </div>
            <input
              type="email"
              autoFocus
              autoCapitalize="none"
              autoCorrect="off"
              spellCheck="false"
              value={emailInput}
              onChange={(e) => setEmailInput(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && tryUnlockAdmin()}
              placeholder="E-mail"
              style={{ width: "100%", padding: "9px 10px", border: `1px solid ${C.line}`, borderRadius: 6, fontSize: 13, marginBottom: 6 }}
            />
            <input
              type={passwordVisible ? "text" : "password"}
              autoCapitalize="none"
              autoCorrect="off"
              spellCheck="false"
              value={passwordInput}
              onChange={(e) => setPasswordInput(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && tryUnlockAdmin()}
              placeholder="Senha"
              style={{ width: "100%", padding: "9px 10px", border: `1px solid ${C.line}`, borderRadius: 6, fontSize: 13, marginBottom: 6 }}
            />
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 11.5, color: C.slate, marginBottom: 8, cursor: "pointer" }}>
              <input type="checkbox" checked={passwordVisible} onChange={(e) => setPasswordVisible(e.target.checked)} />
              mostrar senha
            </label>
            {passwordError && <div style={{ color: C.signal, fontSize: 12, marginBottom: 8 }}>E-mail ou senha incorretos.</div>}
            <button
              onClick={tryUnlockAdmin}
              style={{ width: "100%", padding: "9px", border: "none", borderRadius: 6, background: C.signal, color: "#fff", fontWeight: 700, fontSize: 13, cursor: "pointer" }}
            >
              Entrar
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

function BrandBlock() {
  const [attempt, setAttempt] = useState(0); // 0 = .png, 1 = .svg, 2 = sem imagem (usa tipografia)
  const sources = ["/logo-sudati.png", "/logo-sudati.svg"];

  if (attempt < sources.length) {
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
        <img
          src={sources[attempt]}
          alt="Sudati"
          onError={() => setAttempt((a) => a + 1)}
          style={{ height: 34, maxWidth: 180, objectFit: "contain", objectPosition: "left center" }}
        />
        <div style={{ fontSize: 11.5, color: "#A8B4BB", letterSpacing: "0.01em" }}>Portal de Gestão de Pessoas</div>
      </div>
    );
  }

  return (
    <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
      <div style={{ width: 3, height: 32, background: C.signal, borderRadius: 2, flexShrink: 0 }} />
      <div>
        <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 21, letterSpacing: "0.02em", textTransform: "uppercase", lineHeight: 1 }}>
          Sudati
        </div>
        <div style={{ fontSize: 11.5, color: "#A8B4BB", marginTop: 4, letterSpacing: "0.01em" }}>
          Portal de Gestão de Pessoas
        </div>
      </div>
    </div>
  );
}

function RolePill({ active, onClick, label }) {
  return (
    <button
      className="qlt-pill"
      onClick={onClick}
      style={{
        padding: "5px 12px",
        borderRadius: 20,
        border: `1px solid ${active ? C.signal : "#3E4A52"}`,
        background: active ? C.signal : "transparent",
        color: "#fff",
        fontSize: 12.5,
        fontWeight: 600,
        cursor: "pointer",
      }}
    >
      {label}
    </button>
  );
}

function StatChip({ label, value, accent }) {
  return (
    <div style={{ textAlign: "right" }}>
      <div style={{ fontSize: 10.5, color: "#A8B4BB", textTransform: "uppercase", letterSpacing: "0.08em" }}>{label}</div>
      <div
        style={{
          fontFamily: F.mono,
          fontWeight: 700,
          fontSize: 20,
          color: accent ? "#8FC4E8" : "#FFFFFF",
        }}
      >
        {value}
      </div>
    </div>
  );
}

function SaveIndicator({ state }) {
  if (state === "idle") return null;
  const label = state === "saving" ? "Salvando…" : state === "saved" ? "Salvo" : "Erro ao salvar";
  return (
    <div
      style={{
        fontSize: 11,
        padding: "4px 10px",
        borderRadius: 20,
        background: state === "error" ? "#5A2B22" : "#2E4A3E",
        color: "#F3F5F6",
        display: "flex",
        alignItems: "center",
        gap: 4,
      }}
    >
      {state === "saved" && <Check size={12} />}
      {label}
    </div>
  );
}

// ---------- Tabs ----------
function TabBar({ tab, setTab, counts, role }) {
  const tabs =
    role === "admin"
      ? [
          { id: "adminhome", label: "Início" },
          { id: "dashboard", label: "Visão geral" },
          { id: "lotacao", label: "Quadro de lotação" },
          { id: "organograma", label: "Organograma" },
          { id: "vagas", label: "Vagas em aberto", badge: counts.vagas },
          { id: "indicadores", label: "Indicadores" },
          { id: "solicitacoes", label: "Solicitações", badge: counts.pendentes },
        ]
      : [
          { id: "gestorhome", label: "Início" },
          { id: "meusetor", label: "Meu setor" },
          { id: "vagasgestor", label: "Vagas do meu setor" },
          { id: "abrirvaga", label: "Abrir vaga" },
          { id: "indicadoresgestor", label: "Indicadores" },
        ];
  return (
    <div style={{ background: C.surface, borderBottom: `1px solid ${C.line}` }}>
      <div style={{ maxWidth: 1200, margin: "0 auto", display: "flex", gap: 4, padding: "0 24px", overflowX: "auto" }} className="qlt-scroll">
        {tabs.map((t) => {
          const active = tab === t.id;
          return (
            <button
              key={t.id}
              className="qlt-tab"
              onClick={() => setTab(t.id)}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                padding: "12px 14px",
                background: "transparent",
                border: "none",
                borderBottom: active ? `3px solid ${C.signal}` : "3px solid transparent",
                color: active ? C.ink : C.slateLight,
                fontWeight: active ? 700 : 500,
                fontSize: 13.5,
                cursor: "pointer",
                whiteSpace: "nowrap",
                fontFamily: F.body,
                transition: "color 0.15s ease, border-color 0.15s ease",
              }}
            >
              {t.label}
              {typeof t.badge === "number" && t.badge > 0 && (
                <span
                  style={{
                    background: C.signal,
                    color: "#fff",
                    borderRadius: 20,
                    fontSize: 10.5,
                    fontFamily: F.mono,
                    fontWeight: 700,
                    padding: "1px 6px",
                    minWidth: 16,
                    textAlign: "center",
                  }}
                >
                  {t.badge}
                </span>
              )}
            </button>
          );
        })}
      </div>
    </div>
  );
}

// ---------- Dashboard ----------
function Dashboard({ totals, byArea }) {
  const areaRows = Object.entries(byArea)
    .map(([area, rows]) => ({
      area,
      aut: rows.reduce((s, r) => s + r.autorizadas, 0),
      ocu: rows.reduce((s, r) => s + r.lotadas + r.terceiro, 0),
      abertas: rows.reduce((s, r) => s + vagasAbertas(r), 0),
    }))
    .sort((a, b) => b.abertas - a.abertas);

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 24 }}>
        <BigStat label="Posições autorizadas" value={totals.aut} />
        <BigStat label="Ocupadas (próprio + terceiro)" value={totals.lot + totals.terc} sub={`${totals.terc} terceirizadas`} />
        <BigStat label="Afastados" value={totals.afast} />
        <BigStat label="Posições sem colaborador (quadro)" value={totals.abertas} accent />
      </div>

      <SectionTitle title="Posições sem colaborador por área (quadro de lotação)" />
      <div style={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)", overflow: "hidden" }}>
        {areaRows.map((r, i) => (
          <div
            key={r.area}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 12,
              padding: "10px 16px",
              borderBottom: i < areaRows.length - 1 ? `1px solid ${C.line}` : "none",
            }}
          >
            <div style={{ width: 190, fontSize: 13, fontWeight: 600, flexShrink: 0 }}>{r.area}</div>
            <div style={{ flex: 1, background: C.bg, borderRadius: 4, height: 8, overflow: "hidden" }}>
              <div
                style={{
                  width: r.aut ? `${Math.min(100, (r.ocu / r.aut) * 100)}%` : "0%",
                  height: "100%",
                  background: r.abertas > 0 ? C.amber : C.moss,
                }}
              />
            </div>
            <div style={{ width: 70, fontFamily: F.mono, fontSize: 12.5, color: C.slate, textAlign: "right" }}>
              {r.ocu}/{r.aut}
            </div>
            <div
              style={{
                width: 30,
                textAlign: "center",
                fontFamily: F.mono,
                fontWeight: 700,
                fontSize: 13,
                color: r.abertas > 0 ? C.signal : C.slateLight,
              }}
            >
              {r.abertas}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function BigStat({ label, value, sub, accent }) {
  return (
    <div style={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)", padding: "14px 16px", borderTop: `3px solid ${accent ? C.signal : C.woodDark}` }}>
      <div style={{ fontSize: 11, color: C.slate, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 4 }}>{label}</div>
      <div style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 28, color: accent ? C.signal : C.ink }}>{value}</div>
      {sub && <div style={{ fontSize: 11.5, color: C.slateLight, marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

function SectionTitle({ title }) {
  return (
    <div
      style={{
        fontFamily: F.display,
        fontWeight: 700,
        fontSize: 12.5,
        textTransform: "uppercase",
        letterSpacing: "0.03em",
        color: C.slate,
        marginBottom: 10,
        marginTop: 4,
      }}
    >
      {title}
    </div>
  );
}

// ---------- Quadro de lotação (table) ----------
function LotacaoTable({ areas, areaFilter, setAreaFilter, search, setSearch, filtered, allData, editingCell, setEditingCell, updateRow, addPosition, removePosition, restorePosition }) {
  const [showForm, setShowForm] = useState(false);
  const [showEncerradas, setShowEncerradas] = useState(false);
  const [novaArea, setNovaArea] = useState("");
  const [novaSubarea, setNovaSubarea] = useState("");
  const [novaFuncao, setNovaFuncao] = useState("");
  const [novaAutorizadas, setNovaAutorizadas] = useState(1);

  const encerradas = useMemo(() => {
    return allData.filter((r) => {
      if (r.ativo !== false) return false;
      const matchesArea = areaFilter === "Todas" || r.area === areaFilter;
      const q = search.trim().toLowerCase();
      const matchesSearch = !q || r.funcao.toLowerCase().includes(q) || r.area.toLowerCase().includes(q);
      return matchesArea && matchesSearch;
    });
  }, [allData, areaFilter, search]);

  const rows = showEncerradas ? [...filtered, ...encerradas] : filtered;

  const podeAdicionar = novaArea.trim() && novaFuncao.trim() && novaAutorizadas >= 0;

  const confirmarNova = () => {
    if (!podeAdicionar) return;
    addPosition({ area: novaArea.trim(), subarea: novaSubarea.trim(), funcao: novaFuncao.trim(), autorizadas: Number(novaAutorizadas) || 0 });
    setNovaArea("");
    setNovaSubarea("");
    setNovaFuncao("");
    setNovaAutorizadas(1);
    setShowForm(false);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
        <Filters areas={areas} areaFilter={areaFilter} setAreaFilter={setAreaFilter} search={search} setSearch={setSearch} />
        <button
          onClick={() => setShowForm((v) => !v)}
          style={{
            padding: "8px 14px",
            borderRadius: 8,
            border: "none",
            background: C.signal,
            color: "#fff",
            fontWeight: 700,
            fontSize: 12.5,
            cursor: "pointer",
          }}
        >
          {showForm ? "Cancelar" : "+ Nova função"}
        </button>
      </div>

      {showForm && (
        <div
          style={{
            background: C.surface,
            border: `1px solid ${C.line}`,
            borderRadius: 8,
            boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)",
            padding: 16,
            marginTop: 10,
            display: "grid",
            gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))",
            gap: 10,
            alignItems: "end",
          }}
        >
          <Field label="Área">
            <input value={novaArea} onChange={(e) => setNovaArea(e.target.value)} placeholder="ex: Manutenção" style={inputStyle} />
          </Field>
          <Field label="Subárea (opcional)">
            <input value={novaSubarea} onChange={(e) => setNovaSubarea(e.target.value)} placeholder="ex: Elétrica" style={inputStyle} />
          </Field>
          <Field label="Função">
            <input value={novaFuncao} onChange={(e) => setNovaFuncao(e.target.value)} placeholder="ex: Auxiliar de Operação" style={inputStyle} />
          </Field>
          <Field label="Vagas autorizadas">
            <input type="number" min="0" value={novaAutorizadas} onChange={(e) => setNovaAutorizadas(e.target.value)} style={inputStyle} />
          </Field>
          <button
            onClick={confirmarNova}
            disabled={!podeAdicionar}
            style={{
              padding: "9px 14px",
              borderRadius: 8,
              border: "none",
              background: podeAdicionar ? C.moss : C.line,
              color: podeAdicionar ? "#fff" : C.slateLight,
              fontWeight: 700,
              fontSize: 12.5,
              cursor: podeAdicionar ? "pointer" : "not-allowed",
              height: 38,
            }}
          >
            Adicionar ao quadro
          </button>
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", margin: "10px 0" }}>
        <div style={{ fontSize: 12, color: C.slate }}>
          {filtered.length} posições ativas{showEncerradas ? ` · ${encerradas.length} encerradas exibidas` : ""} · clique em qualquer célula (Área, Subárea, Função, Lotadas, Afastados, Terceiro, Motivo, Observação) para editar
        </div>
        <label style={{ fontSize: 12, color: C.slate, display: "flex", alignItems: "center", gap: 6, cursor: "pointer", whiteSpace: "nowrap" }}>
          <input type="checkbox" checked={showEncerradas} onChange={(e) => setShowEncerradas(e.target.checked)} />
          mostrar encerradas
        </label>
      </div>

      <div style={{ overflowX: "auto", background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)" }} className="qlt-scroll">
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: C.bg, textAlign: "left" }}>
              {["Área", "Subárea", "Função", "Autorizadas", "Lotadas", "Afastados", "Terceiro", "Vagas", "Motivo", "Observação", ""].map((h) => (
                <th key={h} style={{ padding: "9px 12px", fontSize: 10.5, textTransform: "uppercase", letterSpacing: "0.04em", color: C.slate, borderBottom: `1px solid ${C.line}`, whiteSpace: "nowrap" }}>
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => {
              const abertas = vagasAbertas(r);
              const inativa = r.ativo === false;
              return (
                <tr key={r.id} className="qlt-row" style={{ borderBottom: `1px solid ${C.line}`, opacity: inativa ? 0.55 : 1 }}>
                  <EditableCell row={r} field="area" editingCell={editingCell} setEditingCell={setEditingCell} updateRow={updateRow} type="text" />
                  <EditableCell row={r} field="subarea" editingCell={editingCell} setEditingCell={setEditingCell} updateRow={updateRow} type="text" placeholder="—" />
                  <EditableCell row={r} field="funcao" editingCell={editingCell} setEditingCell={setEditingCell} updateRow={updateRow} type="text" />
                  <td style={{ padding: "8px 12px", fontFamily: F.mono, textAlign: "center" }}>{r.autorizadas}</td>
                  <EditableCell row={r} field="lotadas" editingCell={editingCell} setEditingCell={setEditingCell} updateRow={updateRow} type="number" />
                  <EditableCell row={r} field="afastados" editingCell={editingCell} setEditingCell={setEditingCell} updateRow={updateRow} type="number" />
                  <EditableCell row={r} field="terceiro" editingCell={editingCell} setEditingCell={setEditingCell} updateRow={updateRow} type="number" />
                  <td style={{ padding: "8px 12px", textAlign: "center" }}>
                    {abertas > 0 ? (
                      <span style={{ background: C.signalBg, color: C.signal, fontFamily: F.mono, fontWeight: 700, fontSize: 12, padding: "2px 8px", borderRadius: 12 }}>{abertas}</span>
                    ) : (
                      <span style={{ color: C.slateLight, fontSize: 12 }}>—</span>
                    )}
                  </td>
                  <EditableCell row={r} field="motivo" editingCell={editingCell} setEditingCell={setEditingCell} updateRow={updateRow} type="text" placeholder="—" />
                  <EditableCell row={r} field="observacaoQuadro" editingCell={editingCell} setEditingCell={setEditingCell} updateRow={updateRow} type="text" placeholder="—" />
                  <td style={{ padding: "8px 12px", textAlign: "center" }}>
                    {inativa ? (
                      <button
                        onClick={() => restorePosition(r.id)}
                        style={{ padding: "4px 10px", borderRadius: 6, border: `1px solid ${C.line}`, background: "transparent", color: C.moss, fontSize: 11.5, fontWeight: 600, cursor: "pointer" }}
                      >
                        Restaurar
                      </button>
                    ) : (
                      <button
                        onClick={() => {
                          if (window.confirm(`Encerrar a função "${r.funcao}" (${r.area})?`)) removePosition(r.id);
                        }}
                        style={{ padding: "4px 10px", borderRadius: 6, border: `1px solid ${C.line}`, background: "transparent", color: C.slateLight, fontSize: 11.5, fontWeight: 600, cursor: "pointer" }}
                      >
                        Encerrar
                      </button>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function EditableCell({ row, field, editingCell, setEditingCell, updateRow, type, placeholder }) {
  const cellId = `${row.id}-${field}`;
  const isEditing = editingCell === cellId;
  const value = row[field];

  if (isEditing) {
    return (
      <td style={{ padding: "4px 8px" }}>
        <input
          autoFocus
          type={type === "number" ? "number" : "text"}
          defaultValue={value}
          onBlur={(e) => {
            const v = type === "number" ? Math.max(0, parseInt(e.target.value) || 0) : e.target.value;
            updateRow(row.id, field, v);
            setEditingCell(null);
          }}
          onKeyDown={(e) => {
            if (e.key === "Enter") e.target.blur();
            if (e.key === "Escape") setEditingCell(null);
          }}
          style={{
            width: type === "number" ? 56 : 160,
            padding: "4px 6px",
            border: `1px solid ${C.wood}`,
            borderRadius: 4,
            fontSize: 13,
            fontFamily: type === "number" ? F.mono : F.body,
          }}
        />
      </td>
    );
  }
  return (
    <td
      onClick={() => setEditingCell(cellId)}
      style={{
        padding: "8px 12px",
        fontFamily: type === "number" ? F.mono : F.body,
        textAlign: type === "number" ? "center" : "left",
        cursor: "pointer",
        color: value ? C.ink : C.slateLight,
        fontStyle: value ? "normal" : "italic",
      }}
      title="Clique para editar"
    >
      {value || placeholder || "0"}
    </td>
  );
}

function Filters({ areas, areaFilter, setAreaFilter, search, setSearch }) {
  return (
    <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
      <div style={{ position: "relative", flex: "1 1 220px", maxWidth: 320 }}>
        <Search size={14} style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: C.slateLight }} />
        <input
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          placeholder="Buscar função ou área…"
          style={{
            width: "100%",
            padding: "8px 10px 8px 30px",
            border: `1px solid ${C.line}`,
            borderRadius: 6,
            fontSize: 13,
            background: C.surface,
          }}
        />
      </div>
      <select
        value={areaFilter}
        onChange={(e) => setAreaFilter(e.target.value)}
        style={{
          padding: "8px 10px",
          border: `1px solid ${C.line}`,
          borderRadius: 6,
          fontSize: 13,
          background: C.surface,
          color: C.ink,
        }}
      >
        <option>Todas</option>
        {areas.map((a) => (
          <option key={a}>{a}</option>
        ))}
      </select>
      {areaFilter !== "Todas" && (
        <button
          onClick={() => setAreaFilter("Todas")}
          style={{ display: "flex", alignItems: "center", gap: 4, background: "none", border: "none", color: C.slate, fontSize: 12.5, cursor: "pointer" }}
        >
          <X size={13} /> limpar filtro
        </button>
      )}
    </div>
  );
}

// ---------- Organograma ----------
function buildAreaTree(rows) {
  const byId = {};
  rows.forEach((r) => {
    byId[r.id] = { ...r, children: [] };
  });
  const roots = [];
  rows.forEach((r) => {
    const node = byId[r.id];
    if (r.reportaPara && byId[r.reportaPara]) {
      byId[r.reportaPara].children.push(node);
    } else {
      roots.push(node);
    }
  });
  return roots;
}

function isDescendant(allRows, ancestorId, candidateId) {
  const directChildren = allRows.filter((r) => r.reportaPara === ancestorId);
  for (const c of directChildren) {
    if (c.id === candidateId) return true;
    if (isDescendant(allRows, c.id, candidateId)) return true;
  }
  return false;
}

function Organograma({ byArea, areas, expandedAreas, toggleArea, updateRow }) {
  const diretoria = byArea["Diretoria"] || [];
  const otherAreas = areas.filter((a) => a !== "Diretoria");
  const allRows = useMemo(() => Object.values(byArea).flat(), [byArea]);

  const handleDrop = (e, targetId) => {
    e.preventDefault();
    const draggedId = Number(e.dataTransfer.getData("text/plain"));
    if (!draggedId || draggedId === targetId) return;
    if (isDescendant(allRows, draggedId, targetId)) return; // evita ciclo
    updateRow(draggedId, "reportaPara", targetId);
  };

  const handleDropRoot = (e) => {
    e.preventDefault();
    const draggedId = Number(e.dataTransfer.getData("text/plain"));
    if (!draggedId) return;
    updateRow(draggedId, "reportaPara", "");
  };

  return (
    <div>
      <div style={{ fontSize: 12.5, color: C.slate, marginBottom: 16, maxWidth: 640 }}>
        Arraste uma função e solte em cima de outra pra definir quem reporta pra quem. Solte na faixa
        "soltar aqui para tirar do chefe" pra deixar a função no topo da área. Caixas laranja indicam
        vaga em aberto.
      </div>

      {/* Diretoria row */}
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", justifyContent: "center", marginBottom: 22 }}>
        {diretoria.map((r) => (
          <OrgBox key={r.id} row={r} top onDrop={(e) => handleDrop(e, r.id)} />
        ))}
      </div>
      <div style={{ display: "flex", justifyContent: "center", marginBottom: 22 }}>
        <div style={{ width: 2, height: 20, background: C.line }} />
      </div>

      {/* Area cards */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(260px, 1fr))", gap: 12 }}>
        {otherAreas.map((area) => {
          const rows = byArea[area];
          const abertas = rows.reduce((s, r) => s + vagasAbertas(r), 0);
          const expanded = !!expandedAreas[area];
          const tree = buildAreaTree(rows);
          return (
            <div key={area} style={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)", overflow: "hidden", borderTop: `3px solid ${abertas > 0 ? C.signal : C.woodDark}` }}>
              <button
                onClick={() => toggleArea(area)}
                style={{
                  width: "100%",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  padding: "10px 12px",
                  background: "none",
                  border: "none",
                  cursor: "pointer",
                  textAlign: "left",
                }}
              >
                <div>
                  <div style={{ fontWeight: 700, fontSize: 13.5 }}>{area}</div>
                  <div style={{ fontSize: 11, color: C.slateLight }}>{rows.length} funções{abertas > 0 ? ` · ${abertas} vaga${abertas > 1 ? "s" : ""}` : ""}</div>
                </div>
                {expanded ? <ChevronDown size={16} color={C.slate} /> : <ChevronRight size={16} color={C.slate} />}
              </button>
              {expanded && (
                <div style={{ borderTop: `1px solid ${C.line}` }}>
                  <div
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={handleDropRoot}
                    style={{ padding: "4px 12px", fontSize: 10, color: C.slateLight, fontStyle: "italic", borderBottom: `1px dashed ${C.line}` }}
                  >
                    soltar aqui para tirar do chefe
                  </div>
                  {tree.map((node) => (
                    <TreeNode key={node.id} node={node} depth={0} onDrop={handleDrop} />
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function TreeNode({ node, depth, onDrop }) {
  const [expanded, setExpanded] = useState(true);
  const abertas = vagasAbertas(node);
  const hasChildren = node.children && node.children.length > 0;

  return (
    <div>
      <div
        draggable
        onDragStart={(e) => e.dataTransfer.setData("text/plain", String(node.id))}
        onDragOver={(e) => e.preventDefault()}
        onDrop={(e) => onDrop(e, node.id)}
        style={{
          display: "flex",
          alignItems: "center",
          gap: 6,
          padding: "7px 12px",
          paddingLeft: 12 + depth * 16,
          borderBottom: `1px solid ${C.bg}`,
          background: abertas > 0 ? C.signalBg : "transparent",
          cursor: "grab",
        }}
      >
        {hasChildren ? (
          <button
            onClick={() => setExpanded((v) => !v)}
            style={{ background: "none", border: "none", padding: 0, cursor: "pointer", display: "flex", color: C.slateLight }}
          >
            {expanded ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
          </button>
        ) : (
          <span style={{ width: 13, display: "inline-block" }} />
        )}
        <div style={{ fontSize: 12, flex: 1 }}>{node.funcao}</div>
        <div style={{ fontFamily: F.mono, fontSize: 11, color: C.slateLight, whiteSpace: "nowrap", marginLeft: 8 }}>
          {node.lotadas + node.terceiro}/{node.autorizadas}
        </div>
        {abertas > 0 && (
          <div style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 11.5, color: C.signal, marginLeft: 8, minWidth: 14, textAlign: "right" }}>{abertas}</div>
        )}
      </div>
      {hasChildren && expanded && (
        <div style={{ borderLeft: `2px solid ${C.line}`, marginLeft: 18 + depth * 16 }}>
          {node.children.map((child) => (
            <TreeNode key={child.id} node={child} depth={depth + 1} onDrop={onDrop} />
          ))}
        </div>
      )}
    </div>
  );
}

function OrgBox({ row, top, onDrop }) {
  const abertas = vagasAbertas(row);
  return (
    <div
      onDragOver={(e) => onDrop && e.preventDefault()}
      onDrop={onDrop}
      style={{
        background: top ? C.woodDark : C.surface,
        color: top ? "#F3F5F6" : C.ink,
        border: `1px solid ${top ? C.woodDark : C.line}`,
        borderRadius: 6,
        padding: "10px 14px",
        minWidth: 150,
        textAlign: "center",
        position: "relative",
      }}
    >
      <div style={{ fontWeight: 700, fontSize: 12.5 }}>{row.funcao}</div>
      <div style={{ fontSize: 10.5, opacity: 0.75, marginTop: 2 }}>{row.area}</div>
      {abertas > 0 && (
        <div
          style={{
            position: "absolute",
            top: -8,
            right: -8,
            background: C.signal,
            color: "#fff",
            borderRadius: 20,
            fontSize: 10,
            fontFamily: F.mono,
            fontWeight: 700,
            padding: "2px 6px",
          }}
        >
          {abertas}
        </div>
      )}
    </div>
  );
}

// ---------- Vagas em aberto ----------
function VagasAbertas({ vagasLog, updateVaga }) {
  const [setorFiltro, setSetorFiltro] = useState("Todos");
  const [busca, setBusca] = useState("");
  const [mostrarFechadas, setMostrarFechadas] = useState(false);

  const setores = useMemo(() => {
    const seen = [];
    vagasLog.forEach((v) => {
      if (v.setor && !seen.includes(v.setor)) seen.push(v.setor);
    });
    return seen.sort();
  }, [vagasLog]);

  const filtradas = useMemo(() => {
    return vagasLog.filter((v) => {
      const ativa = v.situacao === "Aberta" || v.situacao === "Congelada";
      if (!mostrarFechadas && !ativa) return false;
      if (setorFiltro !== "Todos" && v.setor !== setorFiltro) return false;
      const q = busca.trim().toLowerCase();
      if (q && !v.vaga.toLowerCase().includes(q) && !v.setor.toLowerCase().includes(q) && !(v.gestor || "").toLowerCase().includes(q)) {
        return false;
      }
      return true;
    });
  }, [vagasLog, setorFiltro, busca, mostrarFechadas]);

  const grouped = useMemo(() => {
    const map = {};
    filtradas.forEach((v) => {
      if (!map[v.setor]) map[v.setor] = [];
      map[v.setor].push(v);
    });
    return map;
  }, [filtradas]);

  const abertasCount = vagasLog.filter((v) => v.situacao === "Aberta" || v.situacao === "Congelada").length;

  return (
    <div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
        <div style={{ position: "relative", flex: "1 1 220px", maxWidth: 320 }}>
          <Search size={14} style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: C.slateLight }} />
          <input
            value={busca}
            onChange={(e) => setBusca(e.target.value)}
            placeholder="Buscar vaga, setor ou gestor…"
            style={{ width: "100%", padding: "8px 10px 8px 30px", border: `1px solid ${C.line}`, borderRadius: 6, fontSize: 13, background: C.surface }}
          />
        </div>
        <select value={setorFiltro} onChange={(e) => setSetorFiltro(e.target.value)} style={{ padding: "8px 10px", border: `1px solid ${C.line}`, borderRadius: 6, fontSize: 13, background: C.surface, color: C.ink }}>
          <option>Todos</option>
          {setores.map((s) => (
            <option key={s}>{s}</option>
          ))}
        </select>
        <label style={{ fontSize: 12.5, color: C.slate, display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
          <input type="checkbox" checked={mostrarFechadas} onChange={(e) => setMostrarFechadas(e.target.checked)} />
          mostrar fechadas/canceladas também
        </label>
      </div>

      <div style={{ fontSize: 12, color: C.slate, margin: "10px 0" }}>
        {abertasCount} vagas ativas no total (base: Controle de Vagas 2026) · {filtradas.length} exibidas com esse filtro
      </div>

      {filtradas.length === 0 && (
        <div style={{ textAlign: "center", padding: "60px 20px", color: C.slateLight }}>
          <Briefcase size={28} style={{ marginBottom: 10, opacity: 0.5 }} />
          <div style={{ fontSize: 14 }}>Nenhuma vaga para esse filtro.</div>
        </div>
      )}

      {Object.entries(grouped).map(([setor, rows]) => (
        <div key={setor} style={{ marginBottom: 18 }}>
          <SectionTitle title={setor} />
          <div style={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)", overflow: "hidden" }}>
            {rows.map((v, i) => {
              const isAtiva = v.situacao === "Aberta" || v.situacao === "Congelada";
              return (
                <div
                  key={v.id}
                  style={{
                    padding: "12px 16px",
                    borderBottom: i < rows.length - 1 ? `1px solid ${C.line}` : "none",
                    opacity: isAtiva ? 1 : 0.65,
                  }}
                >
                  <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                    <div style={{ flex: "1 1 200px" }}>
                      <div style={{ fontWeight: 600, fontSize: 13.5 }}>{v.vaga}</div>
                      {v.motivo && <div style={{ fontSize: 11.5, color: C.slateLight, marginTop: 1 }}>{v.motivo}</div>}
                      {v.gestor && <div style={{ fontSize: 11, color: C.slateLight, marginTop: 1 }}>gestor: {v.gestor} · {v.mes}</div>}
                    </div>
                    <select
                      value={v.situacao}
                      onChange={(e) => updateVaga(v.id, "situacao", e.target.value)}
                      style={{
                        padding: "6px 10px",
                        borderRadius: 6,
                        border: `1px solid ${C.line}`,
                        fontSize: 12.5,
                        fontWeight: 700,
                        background: v.situacao === "Aberta" ? C.signalBg : v.situacao === "Congelada" ? C.amberBg : v.situacao === "Fechada" ? C.mossBg : C.bg,
                        color: v.situacao === "Aberta" ? C.signal : v.situacao === "Congelada" ? C.amber : v.situacao === "Fechada" ? C.moss : C.slateLight,
                      }}
                    >
                      {["Aberta", "Congelada", "Fechada", "Cancelada"].map((s) => (
                        <option key={s} value={s}>
                          {s}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns: "repeat(auto-fit, minmax(120px, 1fr))",
                      gap: 8,
                      marginTop: 10,
                      paddingTop: 10,
                      borderTop: `1px dashed ${C.line}`,
                    }}
                  >
                    <MiniField label="Gestor responsável">
                      <input value={v.gestor || ""} onChange={(e) => updateVaga(v.id, "gestor", e.target.value)} style={miniInputStyle} />
                    </MiniField>
                    <MiniField label="Abertura">
                      <input type="date" value={v.dataAbertura || ""} onChange={(e) => updateVaga(v.id, "dataAbertura", e.target.value)} style={miniInputStyle} />
                    </MiniField>
                    <MiniField label="Fechamento">
                      <input type="date" value={v.dataFechamento || ""} onChange={(e) => updateVaga(v.id, "dataFechamento", e.target.value)} style={miniInputStyle} />
                    </MiniField>
                    <MiniField label="Início do contratado">
                      <input type="date" value={v.dataIntegracao || ""} onChange={(e) => updateVaga(v.id, "dataIntegracao", e.target.value)} style={miniInputStyle} />
                    </MiniField>
                    <MiniField label="Entrevistados">
                      <input
                        type="number"
                        min="0"
                        value={v.entrevistados || 0}
                        onChange={(e) => updateVaga(v.id, "entrevistados", Math.max(0, parseInt(e.target.value) || 0))}
                        style={miniInputStyle}
                      />
                    </MiniField>
                    <MiniField label="Mulheres entrevistadas">
                      <input
                        type="number"
                        min="0"
                        value={v.mulheres || 0}
                        onChange={(e) => updateVaga(v.id, "mulheres", Math.max(0, parseInt(e.target.value) || 0))}
                        style={miniInputStyle}
                      />
                    </MiniField>
                    <MiniField label="Recrutamento (I/E)">
                      <select value={v.recrutamento || ""} onChange={(e) => updateVaga(v.id, "recrutamento", e.target.value)} style={miniInputStyle}>
                        <option value="">—</option>
                        <option value="RI">Interno</option>
                        <option value="RE">Externo</option>
                        <option value="RI/RE">Interno + Externo</option>
                      </select>
                    </MiniField>
                    <MiniField label="Nome do aprovado">
                      <input value={v.nomeAprovado || ""} onChange={(e) => updateVaga(v.id, "nomeAprovado", e.target.value)} style={miniInputStyle} />
                    </MiniField>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      ))}
    </div>
  );
}

function MiniField({ label, children }) {
  return (
    <div>
      <div style={{ fontSize: 10, color: C.slateLight, textTransform: "uppercase", letterSpacing: "0.03em", marginBottom: 3 }}>{label}</div>
      {children}
    </div>
  );
}

const miniInputStyle = {
  width: "100%",
  padding: "5px 7px",
  border: `1px solid ${C.line}`,
  borderRadius: 5,
  fontSize: 12,
  background: C.bg,
  color: C.ink,
  fontFamily: F.mono,
};

function FuncaoRow({ row }) {
  const abertas = vagasAbertas(row);
  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        padding: "7px 12px",
        borderBottom: `1px solid ${C.bg}`,
        background: abertas > 0 ? C.signalBg : "transparent",
      }}
    >
      <div style={{ fontSize: 12, flex: 1 }}>{row.funcao}</div>
      <div style={{ fontFamily: F.mono, fontSize: 11, color: C.slateLight, whiteSpace: "nowrap", marginLeft: 8 }}>
        {row.lotadas + row.terceiro}/{row.autorizadas}
      </div>
      {abertas > 0 && (
        <div style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 11.5, color: C.signal, marginLeft: 8, minWidth: 14, textAlign: "right" }}>{abertas}</div>
      )}
    </div>
  );
}

// ---------- Gestor: meu setor ----------
// ---------- Gestor: tela inicial (quadros) ----------
// ---------- Admin: tela inicial (quadros) ----------
function AdminHome({ setTab, vagasCount, pendentesCount }) {
  const cards = [
    {
      id: "dashboard",
      icon: LayoutGrid,
      title: "Visão geral",
      desc: "Ocupação, posições sem colaborador e resumo por área.",
    },
    {
      id: "lotacao",
      icon: Users,
      title: "Quadro de lotação",
      desc: "Edite cargos, áreas e a lotação de cada função.",
    },
    {
      id: "organograma",
      icon: TrendingUp,
      title: "Organograma",
      desc: "Monte a árvore de quem reporta pra quem, arrastando as funções.",
    },
    {
      id: "vagas",
      icon: Briefcase,
      title: "Vagas em aberto",
      desc: "Acompanhe cada processo seletivo em andamento.",
      badge: vagasCount > 0 ? vagasCount : null,
    },
    {
      id: "indicadores",
      icon: BarChart3,
      title: "Indicadores",
      desc: "Tempo de fechamento, entrevistas e recrutamento da empresa toda.",
    },
    {
      id: "solicitacoes",
      icon: FilePlus2,
      title: "Solicitações",
      desc: "Aprove ou recuse pedidos de abertura de vaga dos gestores.",
      badge: pendentesCount > 0 ? pendentesCount : null,
    },
  ];

  return (
    <div>
      <div style={{ marginBottom: 22 }}>
        <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 20, color: C.ink }}>Painel do RH</div>
        <div style={{ fontSize: 13, color: C.slate, marginTop: 4 }}>Escolha o que você quer acessar.</div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14 }}>
        {cards.map((c) => (
          <button
            key={c.id}
            onClick={() => setTab(c.id)}
            className="qlt-card-btn"
            style={{
              textAlign: "left",
              background: C.surface,
              border: `1px solid ${C.line}`,
              borderRadius: 12,
              boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)",
              padding: 20,
              cursor: "pointer",
              position: "relative",
              display: "flex",
              flexDirection: "column",
              gap: 10,
            }}
          >
            {c.badge && (
              <span
                style={{
                  position: "absolute",
                  top: 14,
                  right: 14,
                  background: C.signal,
                  color: "#fff",
                  borderRadius: 20,
                  fontSize: 11,
                  fontFamily: F.mono,
                  fontWeight: 700,
                  padding: "2px 8px",
                }}
              >
                {c.badge}
              </span>
            )}
            <div
              style={{
                width: 40,
                height: 40,
                borderRadius: 10,
                background: C.bg,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
              }}
            >
              <c.icon size={20} color={C.wood} />
            </div>
            <div style={{ fontWeight: 700, fontSize: 15, color: C.ink }}>{c.title}</div>
            <div style={{ fontSize: 12.5, color: C.slate, lineHeight: 1.4 }}>{c.desc}</div>
            <div style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 12, color: C.signal, fontWeight: 600, marginTop: 4 }}>
              Acessar <ArrowRight size={13} />
            </div>
          </button>
        ))}
      </div>
    </div>
  );
}

function GestorHome({ area, setTab, vagasCount }) {
  const cards = [
    {
      id: "meusetor",
      icon: Users,
      title: "Meu setor",
      desc: "Veja o quadro de posições e a ocupação do seu time.",
    },
    {
      id: "vagasgestor",
      icon: Briefcase,
      title: "Vagas do meu setor",
      desc: "Acompanhe o status das vagas em processo seletivo.",
      badge: vagasCount > 0 ? vagasCount : null,
    },
    {
      id: "abrirvaga",
      icon: FilePlus2,
      title: "Abrir vaga",
      desc: "Solicite uma reposição ou um aumento de quadro.",
    },
    {
      id: "indicadoresgestor",
      icon: BarChart3,
      title: "Indicadores",
      desc: "Tempo de fechamento, entrevistas e recrutamento do seu setor.",
    },
  ];

  return (
    <div>
      <div style={{ marginBottom: 22 }}>
        <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 20, color: C.ink }}>Olá! O que você quer fazer hoje?</div>
        <div style={{ fontSize: 13, color: C.slate, marginTop: 4 }}>
          Setor selecionado: <strong style={{ color: C.ink }}>{area}</strong>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14 }}>
        {cards.map((c) => (
          <button
            key={c.id}
            onClick={() => setTab(c.id)}
            className="qlt-card-btn"
            style={{
              textAlign: "left",
              background: C.surface,
              border: `1px solid ${C.line}`,
              borderRadius: 12,
              boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)",
              padding: 20,
              cursor: "pointer",
              position: "relative",
              display: "flex",
              flexDirection: "column",
              gap: 10,
            }}
          >
            {c.badge && (
              <span
                style={{
                  position: "absolute",
                  top: 14,
                  right: 14,
                  background: C.signal,
                  color: "#fff",
                  borderRadius: 20,
                  fontSize: 11,
                  fontFamily: F.mono,
                  fontWeight: 700,
                  padding: "2px 8px",
                }}
              >
                {c.badge}
              </span>
            )}
            <div
              style={{
                width: 40,
                height: 40,
                borderRadius: 10,
                background: C.bg,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
              }}
            >
              <c.icon size={20} color={C.wood} />
            </div>
            <div style={{ fontWeight: 700, fontSize: 15, color: C.ink }}>{c.title}</div>
            <div style={{ fontSize: 12.5, color: C.slate, lineHeight: 1.4 }}>{c.desc}</div>
            <div style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 12, color: C.signal, fontWeight: 600, marginTop: 4 }}>
              Acessar <ArrowRight size={13} />
            </div>
          </button>
        ))}
      </div>
    </div>
  );
}

function MeuSetor({ area, rows }) {
  const aut = rows.reduce((s, r) => s + r.autorizadas, 0);
  const ocu = rows.reduce((s, r) => s + r.lotadas + r.terceiro, 0);
  const abertas = rows.reduce((s, r) => s + vagasAbertas(r), 0);
  return (
    <div>
      <SectionTitle title={area} />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 20 }}>
        <BigStat label="Posições autorizadas" value={aut} />
        <BigStat label="Ocupadas" value={ocu} />
        <BigStat label="Vagas em aberto" value={abertas} accent />
      </div>
      <div style={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)", overflow: "hidden" }}>
        {rows.map((r, i) => (
          <FuncaoRow key={r.id} row={r} />
        ))}
      </div>
    </div>
  );
}

// ---------- Gestor: vagas do meu setor ----------
function VagasGestor({ rows }) {
  if (rows.length === 0) {
    return (
      <div style={{ textAlign: "center", padding: "60px 20px", color: C.slateLight }}>
        <Briefcase size={28} style={{ marginBottom: 10, opacity: 0.5 }} />
        <div style={{ fontSize: 14 }}>Nenhuma vaga em aberto no seu setor no momento.</div>
      </div>
    );
  }
  return (
    <div>
      <div style={{ fontSize: 12, color: C.slate, marginBottom: 10 }}>
        Acompanhamento do processo seletivo — o status é atualizado pelo RH.
      </div>
      <div style={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)", overflow: "hidden" }}>
        {rows.map((r, i) => (
          <div
            key={r.id}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 12,
              padding: "12px 16px",
              borderBottom: i < rows.length - 1 ? `1px solid ${C.line}` : "none",
              flexWrap: "wrap",
            }}
          >
            <div style={{ flex: "1 1 200px" }}>
              <div style={{ fontWeight: 600, fontSize: 13.5 }}>{r.funcao}</div>
              {r.motivo && <div style={{ fontSize: 11.5, color: C.slateLight, marginTop: 1 }}>{r.motivo}</div>}
            </div>
            <div
              style={{
                fontFamily: F.mono,
                fontWeight: 700,
                fontSize: 13,
                color: C.signal,
                background: C.signalBg,
                borderRadius: 12,
                padding: "3px 10px",
                whiteSpace: "nowrap",
              }}
            >
              {vagasAbertas(r)} vaga{vagasAbertas(r) > 1 ? "s" : ""}
            </div>
            <div
              style={{
                padding: "6px 10px",
                borderRadius: 6,
                fontSize: 12.5,
                background: r.status ? C.mossBg : C.bg,
                color: r.status ? C.moss : C.slateLight,
                fontWeight: 600,
              }}
            >
              {r.status || "Aguardando RH definir status"}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ---------- Gestor: abrir vaga ----------
function AbrirVaga({ area, rows, requests, submitRequest, gestorNome, setGestorNome }) {
  const [funcaoSel, setFuncaoSel] = useState("");
  const [funcaoNova, setFuncaoNova] = useState("");
  const [usarNova, setUsarNova] = useState(false);
  const [tipo, setTipo] = useState(TIPO_SOLICITACAO[0]);
  const [colaboradorSubstituido, setColaboradorSubstituido] = useState("");
  const [justificativa, setJustificativa] = useState("");
  const [dataAbertura, setDataAbertura] = useState("");
  const [confirmado, setConfirmado] = useState(false);

  const isReposicao = tipo === TIPO_SOLICITACAO[0];
  const funcaoFinal = usarNova ? funcaoNova.trim() : funcaoSel;
  const podeEnviar =
    gestorNome.trim() &&
    funcaoFinal &&
    dataAbertura &&
    (isReposicao ? colaboradorSubstituido.trim() : justificativa.trim().length > 4);

  const enviar = () => {
    if (!podeEnviar) return;
    submitRequest({
      area,
      funcao: funcaoFinal,
      tipo,
      colaboradorSubstituido: isReposicao ? colaboradorSubstituido.trim() : "",
      justificativa: isReposicao ? "" : justificativa.trim(),
      solicitante: gestorNome.trim(),
      dataAbertura,
    });
    setConfirmado(true);
    setFuncaoSel("");
    setFuncaoNova("");
    setUsarNova(false);
    setColaboradorSubstituido("");
    setJustificativa("");
    setDataAbertura("");
    setTimeout(() => setConfirmado(false), 2500);
  };

  return (
    <div style={{ maxWidth: 560 }}>
      <SectionTitle title={`Solicitar abertura de vaga — ${area}`} />
      <div style={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)", padding: 18 }}>
        <Field label="Nome do gestor">
          <input value={gestorNome} onChange={(e) => setGestorNome(e.target.value)} placeholder="Seu nome" style={inputStyle} />
        </Field>

        <Field label="Função">
          {!usarNova ? (
            <select
              value={funcaoSel}
              onChange={(e) => setFuncaoSel(e.target.value)}
              style={inputStyle}
            >
              <option value="">Selecione a função…</option>
              {rows.map((r) => (
                <option key={r.id} value={r.funcao}>
                  {r.funcao}
                </option>
              ))}
            </select>
          ) : (
            <input value={funcaoNova} onChange={(e) => setFuncaoNova(e.target.value)} placeholder="Nome da nova função" style={inputStyle} />
          )}
          <button
            onClick={() => setUsarNova((v) => !v)}
            style={{ background: "none", border: "none", color: C.wood, fontSize: 12, cursor: "pointer", marginTop: 6, padding: 0 }}
          >
            {usarNova ? "usar uma função já existente" : "é uma função que ainda não existe no quadro?"}
          </button>
        </Field>

        <Field label="Tipo de solicitação">
          <select value={tipo} onChange={(e) => setTipo(e.target.value)} style={inputStyle}>
            {TIPO_SOLICITACAO.map((t) => (
              <option key={t} value={t}>
                {t}
              </option>
            ))}
          </select>
        </Field>

        <Field label="Data de abertura da vaga">
          <input
            type="date"
            value={dataAbertura}
            onChange={(e) => setDataAbertura(e.target.value)}
            style={inputStyle}
          />
        </Field>

        {isReposicao ? (
          <Field label="Nome do colaborador a ser substituído">
            <input
              value={colaboradorSubstituido}
              onChange={(e) => setColaboradorSubstituido(e.target.value)}
              placeholder="Nome de quem saiu da posição"
              style={inputStyle}
            />
          </Field>
        ) : (
          <Field label="Justificativa">
            <textarea
              value={justificativa}
              onChange={(e) => setJustificativa(e.target.value)}
              placeholder="Explique o motivo do aumento de quadro…"
              rows={3}
              style={{ ...inputStyle, resize: "vertical", fontFamily: F.body }}
            />
          </Field>
        )}

        <button
          onClick={enviar}
          disabled={!podeEnviar}
          style={{
            marginTop: 6,
            padding: "10px 18px",
            borderRadius: 6,
            border: "none",
            background: podeEnviar ? C.signal : C.line,
            color: podeEnviar ? "#fff" : C.slateLight,
            fontWeight: 700,
            fontSize: 13.5,
            cursor: podeEnviar ? "pointer" : "not-allowed",
          }}
        >
          Enviar para aprovação do RH
        </button>
        {confirmado && <div style={{ marginTop: 8, fontSize: 12.5, color: C.moss, fontWeight: 600 }}>Solicitação enviada — o RH foi notificado.</div>}
      </div>

      {requests.length > 0 && (
        <div style={{ marginTop: 24 }}>
          <SectionTitle title="Minhas solicitações" />
          <div style={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)", overflow: "hidden" }}>
            {requests
              .slice()
              .sort((a, b) => b.id - a.id)
              .map((r, i, arr) => (
                <div key={r.id} style={{ padding: "10px 16px", borderBottom: i < arr.length - 1 ? `1px solid ${C.line}` : "none" }}>
                  <div style={{ fontWeight: 600, fontSize: 13 }}>{r.funcao}</div>
                  <div style={{ fontSize: 11.5, color: C.slateLight, marginTop: 1 }}>{r.tipo}</div>
                  {r.colaboradorSubstituido && <div style={{ fontSize: 11.5, color: C.slateLight, marginTop: 1 }}>substitui: {r.colaboradorSubstituido}</div>}
                  <StatusBadge status={r.status} />
                </div>
              ))}
          </div>
        </div>
      )}
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <div style={{ fontSize: 11.5, color: C.slate, textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 5, fontWeight: 600 }}>{label}</div>
      {children}
    </div>
  );
}

const inputStyle = {
  width: "100%",
  padding: "9px 10px",
  border: `1px solid ${C.line}`,
  borderRadius: 6,
  fontSize: 13,
  background: C.bg,
  color: C.ink,
};

function StatusBadge({ status }) {
  const color = status === "Aprovada" ? C.moss : status === "Recusada" ? C.signal : C.amber;
  const bg = status === "Aprovada" ? C.mossBg : status === "Recusada" ? C.signalBg : C.amberBg;
  return (
    <span style={{ display: "inline-block", marginTop: 4, fontSize: 11, fontWeight: 700, color, background: bg, borderRadius: 12, padding: "2px 9px" }}>
      {status}
    </span>
  );
}

// ---------- Admin: solicitações ----------
function Solicitacoes({ requests, resolveRequest }) {
  if (requests.length === 0) {
    return (
      <div style={{ textAlign: "center", padding: "60px 20px", color: C.slateLight }}>
        <Briefcase size={28} style={{ marginBottom: 10, opacity: 0.5 }} />
        <div style={{ fontSize: 14 }}>Nenhuma solicitação de vaga enviada pelos gestores ainda.</div>
      </div>
    );
  }
  const pendentes = requests.filter((r) => r.status === "Pendente");
  const resolvidas = requests.filter((r) => r.status !== "Pendente");

  return (
    <div>
      <SectionTitle title={`Pendentes (${pendentes.length})`} />
      {pendentes.length === 0 && <div style={{ fontSize: 12.5, color: C.slateLight, marginBottom: 20 }}>Nenhuma solicitação pendente.</div>}
      {pendentes.length > 0 && (
        <div style={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)", overflow: "hidden", marginBottom: 24 }}>
          {pendentes.map((r, i) => (
            <div
              key={r.id}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 12,
                padding: "12px 16px",
                borderBottom: i < pendentes.length - 1 ? `1px solid ${C.line}` : "none",
                flexWrap: "wrap",
              }}
            >
              <div style={{ flex: "1 1 260px" }}>
                <div style={{ fontWeight: 700, fontSize: 13.5 }}>
                  {r.funcao} <span style={{ fontWeight: 400, color: C.slateLight }}>· {r.area}</span>
                </div>
                <div style={{ fontSize: 11.5, color: C.slate, marginTop: 2 }}>
                  {r.tipo}
                  {r.solicitante ? ` · solicitado por ${r.solicitante}` : ""}
                </div>
                <div style={{ fontSize: 12, color: C.ink, marginTop: 4, fontStyle: "italic" }}>
                  {r.colaboradorSubstituido ? `substitui: ${r.colaboradorSubstituido}` : `"${r.justificativa}"`}
                </div>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <button
                  onClick={() => resolveRequest(r.id, "Aprovada")}
                  style={{ padding: "7px 14px", borderRadius: 6, border: "none", background: C.moss, color: "#fff", fontWeight: 700, fontSize: 12.5, cursor: "pointer" }}
                >
                  Aprovar
                </button>
                <button
                  onClick={() => resolveRequest(r.id, "Recusada")}
                  style={{ padding: "7px 14px", borderRadius: 6, border: `1px solid ${C.line}`, background: "transparent", color: C.slate, fontWeight: 600, fontSize: 12.5, cursor: "pointer" }}
                >
                  Recusar
                </button>
              </div>
            </div>
          ))}
        </div>
      )}

      {resolvidas.length > 0 && (
        <div>
          <SectionTitle title="Histórico" />
          <div style={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)", overflow: "hidden" }}>
            {resolvidas
              .slice()
              .sort((a, b) => b.id - a.id)
              .map((r, i, arr) => (
                <div key={r.id} style={{ padding: "10px 16px", borderBottom: i < arr.length - 1 ? `1px solid ${C.line}` : "none", display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                  <div style={{ flex: "1 1 240px", fontSize: 13 }}>
                    <strong>{r.funcao}</strong> · {r.area}
                  </div>
                  <StatusBadge status={r.status} />
                </div>
              ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ---------- Indicadores ----------
function diffDias(a, b) {
  const d1 = new Date(a);
  const d2 = new Date(b);
  if (isNaN(d1) || isNaN(d2)) return null;
  return Math.round((d2 - d1) / (1000 * 60 * 60 * 24));
}

const MESES_ORDEM = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

function Indicadores({ vagasLog }) {
  const fechadas = vagasLog.filter((v) => v.situacao === "Fechada");
  const abertas = vagasLog.filter((v) => v.situacao === "Aberta");
  const canceladas = vagasLog.filter((v) => v.situacao === "Cancelada");

  const temposFechamento = fechadas
    .map((v) => (v.tempoFechamento > 0 ? v.tempoFechamento : diffDias(v.dataAbertura, v.dataFechamento)))
    .filter((d) => d !== null && d >= 0);
  const tempoMedioFechamento = temposFechamento.length
    ? Math.round(temposFechamento.reduce((s, d) => s + d, 0) / temposFechamento.length)
    : null;

  const temposAdmissao = fechadas.map((v) => v.tempoAdmissao).filter((d) => d && d > 0);
  const tempoMedioAdmissao = temposAdmissao.length
    ? Math.round(temposAdmissao.reduce((s, d) => s + d, 0) / temposAdmissao.length)
    : null;

  const totalEntrevistados = vagasLog.reduce((s, v) => s + (v.entrevistados || 0), 0);
  const totalMulheres = vagasLog.reduce((s, v) => s + (v.mulheres || 0), 0);
  const pctMulheres = totalEntrevistados > 0 ? Math.round((totalMulheres / totalEntrevistados) * 100) : null;

  const interno = fechadas.filter((v) => (v.recrutamento || "").includes("RI")).length;
  const externo = fechadas.filter((v) => (v.recrutamento || "") === "RE").length;

  const porMes = useMemo(() => {
    const map = {};
    vagasLog.forEach((v) => {
      if (!v.mes) return;
      if (!map[v.mes]) map[v.mes] = { mes: v.mes, abertas: 0, fechadas: 0, tempos: [] };
      if (v.situacao === "Aberta") map[v.mes].abertas += 1;
      if (v.situacao === "Fechada") {
        map[v.mes].fechadas += 1;
        const t = v.tempoFechamento > 0 ? v.tempoFechamento : diffDias(v.dataAbertura, v.dataFechamento);
        if (t !== null && t >= 0) map[v.mes].tempos.push(t);
      }
    });
    return Object.values(map)
      .sort((a, b) => MESES_ORDEM.indexOf(a.mes) - MESES_ORDEM.indexOf(b.mes))
      .map((m) => ({
        mes: m.mes.slice(0, 3),
        Abertas: m.abertas,
        Fechadas: m.fechadas,
        "Tempo médio": m.tempos.length ? Math.round(m.tempos.reduce((s, d) => s + d, 0) / m.tempos.length) : null,
      }));
  }, [vagasLog]);

  const porSetorTempo = useMemo(() => {
    const map = {};
    fechadas.forEach((v) => {
      const t = v.tempoFechamento > 0 ? v.tempoFechamento : diffDias(v.dataAbertura, v.dataFechamento);
      if (t === null || t < 0) return;
      if (!map[v.setor]) map[v.setor] = [];
      map[v.setor].push(t);
    });
    return Object.entries(map)
      .map(([setor, tempos]) => ({
        setor,
        tempoMedio: Math.round(tempos.reduce((s, d) => s + d, 0) / tempos.length),
        qtd: tempos.length,
      }))
      .sort((a, b) => b.tempoMedio - a.tempoMedio);
  }, [fechadas]);

  const recrutamentoPie = [
    { name: "Externo", value: externo },
    { name: "Interno", value: interno },
  ].filter((d) => d.value > 0);

  if (vagasLog.length === 0) {
    return (
      <div style={{ textAlign: "center", padding: "60px 20px", color: C.slateLight }}>
        <Briefcase size={28} style={{ marginBottom: 10, opacity: 0.5 }} />
        <div style={{ fontSize: 14 }}>Ainda não há histórico de vagas registrado.</div>
      </div>
    );
  }

  const tooltipStyle = {
    background: C.woodDark,
    border: "none",
    borderRadius: 6,
    color: "#F3F5F6",
    fontSize: 12,
    fontFamily: F.body,
  };
  const axisStyle = { fontSize: 11, fill: C.slateLight, fontFamily: F.body };

  return (
    <div>
      <div style={{ fontSize: 12.5, color: C.slate, marginBottom: 16, maxWidth: 640 }}>
        Calculado a partir do Controle de Vagas — abertura, fechamento, entrevistas e admissão de cada processo.
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10, marginBottom: 20 }}>
        <KpiCard label="Vagas abertas hoje" value={abertas.length} accent />
        <KpiCard label="Fechadas (histórico)" value={fechadas.length} />
        <KpiCard label="Canceladas" value={canceladas.length} />
        <KpiCard label="Tempo médio p/ fechar" value={tempoMedioFechamento ?? "—"} suffix=" dias" />
        <KpiCard label="Tempo médio p/ admissão" value={tempoMedioAdmissao ?? "—"} suffix=" dias" />
        <KpiCard label="Total de entrevistas" value={totalEntrevistados} />
        <KpiCard label="Mulheres entrevistadas" value={pctMulheres !== null ? pctMulheres : "—"} suffix={pctMulheres !== null ? "%" : ""} />
        <KpiCard label="Recrutamento I / E" value={`${interno} / ${externo}`} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr)", gap: 12, marginBottom: 12 }}>
        <ChartCard title="Recrutamento interno x externo">
          <ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie
                data={recrutamentoPie}
                dataKey="value"
                nameKey="name"
                cx="50%"
                cy="50%"
                innerRadius={45}
                outerRadius={75}
                paddingAngle={3}
                label={(entry) => `${entry.name}: ${entry.value}`}
                labelLine={false}
              >
                {recrutamentoPie.map((entry, i) => (
                  <Cell key={entry.name} fill={entry.name === "Interno" ? C.moss : C.wood} />
                ))}
              </Pie>
              <Tooltip contentStyle={tooltipStyle} />
            </PieChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Tempo médio de fechamento por mês (dias)">
          <ResponsiveContainer width="100%" height={220}>
            <LineChart data={porMes} margin={{ top: 8, right: 16, left: -16, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
              <XAxis dataKey="mes" tick={axisStyle} axisLine={{ stroke: C.line }} tickLine={false} />
              <YAxis tick={axisStyle} axisLine={false} tickLine={false} allowDecimals={false} />
              <Tooltip contentStyle={tooltipStyle} />
              <Line type="monotone" dataKey="Tempo médio" stroke={C.signal} strokeWidth={2.5} dot={{ r: 4, fill: C.signal }} connectNulls />
            </LineChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>

      <ChartCard title="Tempo médio de fechamento de vaga por setor (dias)">
        <ResponsiveContainer width="100%" height={Math.max(180, porSetorTempo.length * 30)}>
          <BarChart data={porSetorTempo} layout="vertical" margin={{ top: 4, right: 24, left: 8, bottom: 4 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.line} horizontal={false} />
            <XAxis type="number" tick={axisStyle} axisLine={false} tickLine={false} allowDecimals={false} />
            <YAxis type="category" dataKey="setor" tick={{ ...axisStyle, fontSize: 11 }} width={140} axisLine={false} tickLine={false} />
            <Tooltip contentStyle={tooltipStyle} cursor={{ fill: C.bg }} formatter={(value, name, props) => [`${value} dias (${props.payload.qtd} vaga${props.payload.qtd > 1 ? "s" : ""})`, "Tempo médio"]} />
            <Bar dataKey="tempoMedio" fill={C.wood} radius={[0, 4, 4, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </ChartCard>
    </div>
  );
}

function KpiCard({ label, value, suffix, accent }) {
  return (
    <div
      style={{
        background: C.surface,
        border: `1px solid ${C.line}`,
        borderRadius: 8,
        boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)",
        padding: "14px 16px",
        borderTop: `3px solid ${accent ? C.signal : C.woodDark}`,
      }}
    >
      <div style={{ fontSize: 10.5, color: C.slate, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>{label}</div>
      <div style={{ fontFamily: F.mono, fontWeight: 800, fontSize: 24, color: accent ? C.signal : C.ink, lineHeight: 1 }}>
        {value}
        {suffix && <span style={{ fontSize: 13, fontWeight: 600, color: C.slateLight }}>{suffix}</span>}
      </div>
    </div>
  );
}

function ChartCard({ title, children, style }) {
  return (
    <div
      style={{
        background: C.surface,
        border: `1px solid ${C.line}`,
        borderRadius: 8,
        boxShadow: "0 1px 2px rgba(42,36,28,0.05), 0 2px 8px rgba(42,36,28,0.06)",
        padding: "14px 16px",
        ...style,
      }}
    >
      <div style={{ fontSize: 12, fontWeight: 700, color: C.slate, textTransform: "uppercase", letterSpacing: "0.03em", marginBottom: 6 }}>{title}</div>
      {children}
    </div>
  );
}
