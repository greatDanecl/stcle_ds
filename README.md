# STCLE – Dashboard Roles CABLU

Dashboard de KPIs para la Directiva del Sindicato de Tripulantes de Cabina LanExpress.

## Estructura

```
├── data/              ← Aquí van los .xlsx de roles publicados
├── src/
│   ├── parser.py      ← Genera dashboard_data.json
│   └── logo_b64.txt   ← Logo codificado (generado una vez)
├── index.html         ← Dashboard (se actualiza automáticamente)
└── .github/workflows/
    └── process_data.yml
```

## Cómo actualizar

1. Sube el nuevo archivo `.xlsx` a la carpeta `data/`
2. El workflow se activa automáticamente
3. El dashboard en GitHub Pages se actualiza en ~2 minutos

## KPIs incluidos

| Código | Nombre | Norma |
|--------|--------|-------|
| KPI-1 | Vuelos en franja nocturna 00:30–05:30 | DAN 121 |
| PSVNC | Pares de Servicios con Vuelo Nocturno Consecutivo + descanso < 10h | DAN 121 |
| KPI-2 | Días de descanso DO+DR | DAN 121 |
| KPI-3 | Vacaciones VC | Ley 20.321 |
| KPI-4 | Licencias SICK/LNP/OOF | Ley 20.321 |
| KPI-5 | Standby B+ASB | Ley 20.321 |
| KPI-6 | Horas de vuelo promedio (block time) | DAN 121 |
| KPI-7 | Socios con capacitación ADM/ASB/HSB/CRM | Operacional |
| KPI-8 | Semáforo de alertas por socio | Sindical |

## GitHub Pages

Activar en **Settings → Pages → Source: GitHub Actions**.
