# Parser de Mermas — Carnes Frías La Villa

## Uso en cada sesión de Claude

1. Descargar el archivo de producción desde Google Drive (ID: `1HN_C9tPFb94tMIF3Krjr2nG2XfxCSLzM`)
2. Guardar el texto en `/tmp/prod_full.txt`
3. Ejecutar `parser_lavilla.py` con el texto cargado
4. Inyectar los records en el HTML del dashboard
5. **VALIDAR**: suma de col K en Excel ≈ total "real" del dashboard

## Validación obligatoria antes de publicar

```
Suma col K (Excel, col "peso lote terminado", 7-abril al último día)
≈ Total real del dashboard + peso productos excluidos (jamones, tajados)
Diferencia aceptable: < 15 ton (productos excluidos del análisis)
```

## Reglas críticas resumidas

| Regla | Detalle |
|-------|---------|
| Fuente real | Col K (col[10]) de Entrega PT — báscula desde 7-abril-2026 |
| Patron subdividen | Una vez por lote-base (Chorizo, Sevillano, Schon Pollo, Schon Cervecero) |
| Patron independientes | Suma por cada fila (Manguera, Mortadela bloque, Salchicha, Schon Económico, Schon Promoción) |
| Override PDN | Solo familias que subdividen con cant_ejecutada ≠ 1 |
| Semáforo verde | patron ≤ real ≤ teórico |
| Semáforo amarillo | real < patron (merma mayor al estándar) |
| Semáforo rojo | real > teórico (revisar — puede ser reproceso) |

## Limitaciones conocidas (pendientes)

- Entregas parciales generan alertas falsas — Javi notifica manualmente
- Reproceso agrega peso al lote sin registrarse en PDN → rojos esperados
- Tajados (Mortadela x250, x450, jamonada) excluidos — fase futura

## Archivos en este repositorio

- `index.html` — dashboard publicado en GitHub Pages
- `parser_lavilla.py` — script de parseo y generación de datos
- `README.md` — este archivo
