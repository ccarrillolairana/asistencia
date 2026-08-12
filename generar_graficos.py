import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

# Configuración de estilo académico
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['axes.edgecolor'] = '#CCCCCC'
plt.rcParams['axes.linewidth'] = 0.8

# 1. GRÁFICO TEMA 1: CRECIMIENTO PAGO QR EN BOLIVIA
fig, ax1 = plt.subplots(figsize=(8, 4.5), dpi=300)

años = ['2020', '2022', '2024', '2026 (Proy.)']
operaciones = [12.4, 85.2, 210.0, 380.0]  # En millones
montos = [1.2, 9.5, 26.0, 48.0] # En miles de millones de Bs

bars = ax1.bar(años, operaciones, color='#1F497D', width=0.45, label='Operaciones (Millones)', zorder=2)
ax1.set_ylabel('N° de Operaciones (Millones de transacciones)', color='#1F497D', fontweight='bold', fontsize=10)
ax1.tick_params(axis='y', labelcolor='#1F497D')
ax1.set_ylim(0, 430)

# Agregar etiquetas sobre las barras
for bar in bars:
    yval = bar.get_height()
    ax1.text(bar.get_x() + bar.get_width()/2.0, yval + 10, f'{yval:.1f} M', ha='center', va='bottom', fontsize=9, fontweight='bold', color='#1F497D')

# Eje secundario para montos
ax2 = ax1.twinx()
line = ax2.plot(años, montos, color='#E36C09', marker='o', linewidth=3, markersize=8, label='Monto Total (Miles de Millones Bs)', zorder=3)
ax2.set_ylabel('Monto Total Transaccionado (Miles de Millones Bs)', color='#E36C09', fontweight='bold', fontsize=10)
ax2.tick_params(axis='y', labelcolor='#E36C09')
ax2.set_ylim(0, 55)
ax2.grid(False)

# Etiquetas sobre línea
for i, txt in enumerate(montos):
    ax2.annotate(f'Bs {txt:.1f}B', (años[i], montos[i]), textcoords="offset points", xytext=(0,10), ha='center', fontsize=9, fontweight='bold', color='#E36C09')

plt.title('Evolución Histórica y Proyección del Pago QR Simple en Bolivia (2020-2026)', fontsize=12, fontweight='bold', pad=15, color='#1F497D')
fig.tight_layout()
plt.savefig('grafico_tema1_estadisticas_qr.png', dpi=300)
plt.close()
print("Gráfico 1 generado: grafico_tema1_estadisticas_qr.png")


# 2. FLUJOGRAMA TEMA 1: PROCESO OPERATIVO QR SIMPLE Y BILLETERA MÓVIL
fig, ax = plt.subplots(figsize=(10, 3.5), dpi=300)
ax.axis('off')

boxes = [
    ("1. Usuario / Cliente", "Abre App Bancaria / Soli / Tigo Money\ny selecciona 'Pago QR'"),
    ("2. Escaneo de Código", "Escanea el código QR del comercio\no vendedor ambulante"),
    ("3. Procesamiento BCB / Asoban", "Switch de pago valida cuenta\ne intercambia datos al instante"),
    ("4. Confirmación Inmediata", "Transferencia en 0 segundos\nNotificación SMS / App a ambos")
]

colors = ['#1F497D', '#2E75B6', '#E36C09', '#385723']

for i, (title, desc) in enumerate(boxes):
    # Dibuja rectángulo box
    rect = patches.FancyBboxPatch((i*2.4 + 0.1, 0.3), 2.0, 1.4, boxstyle="round,pad=0.1", ec="none", fc=colors[i])
    ax.add_patch(rect)
    ax.text(i*2.4 + 1.1, 1.3, title, color='white', weight='bold', fontsize=10, ha='center', va='center')
    ax.text(i*2.4 + 1.1, 0.8, desc, color='white', fontsize=8, ha='center', va='center', multialignment='center')
    
    # Dibuja flecha entre cuadros
    if i < len(boxes) - 1:
        ax.annotate('', xy=(i*2.4 + 2.35, 1.0), xytext=(i*2.4 + 2.15, 1.0),
                    arrowprops=dict(arrowstyle="->", color='#333333', lw=2.5))

ax.set_xlim(0, 9.6)
ax.set_ylim(0, 2.0)
plt.title('Flujograma Operativo: Sistema de Pago e Inclusión Financiera QR Simple en Bolivia', fontsize=11, fontweight='bold', pad=10, color='#1F497D')
fig.tight_layout()
plt.savefig('grafico_tema1_flujograma.png', dpi=300)
plt.close()
print("Flujograma 1 generado: grafico_tema1_flujograma.png")


# 3. GRÁFICO TEMA 2: COMPARATIVA DE COSTOS VEHÍCULO A GASOLINA VS ELÉCTRICO QUANTUM
fig, ax = plt.subplots(figsize=(8, 4.5), dpi=300)

categorias = ['Costo por 100 km (Bs)', 'Mantenimiento Anual (Bs)']
gasolina = [42.0, 2500.0]
electrico = [11.5, 600.0]

x = np.arange(len(categorias))
width = 0.35

rects1 = ax.bar(x - width/2, gasolina, width, label='Auto a Gasolina Subvencionada (1.2L)', color='#C00000')
rects2 = ax.bar(x + width/2, electrico, width, label='Auto Eléctrico Quantum E4 (220V)', color='#385723')

ax.set_ylabel('Monto en Bolivianos (Bs)', fontweight='bold', fontsize=10, color='#1F497D')
ax.set_title('Comparativa de Costo Operativo y Mantenimiento en Bolivia (Bs)', fontsize=12, fontweight='bold', pad=15, color='#1F497D')
ax.set_xticks(x)
ax.set_xticklabels(categorias, fontweight='bold', fontsize=10)
ax.legend(frameon=True, facecolor='white', edgecolor='#CCCCCC')

# Añadir valores y porcentaje de ahorro
ax.text(x[0] - width/2, gasolina[0] + 1, 'Bs 42.00', ha='center', fontweight='bold', color='#C00000')
ax.text(x[0] + width/2, electrico[0] + 1, 'Bs 11.50\n(72.6% Ahorro)', ha='center', fontweight='bold', color='#385723')

ax.text(x[1] - width/2, gasolina[1] + 50, 'Bs 2.500', ha='center', fontweight='bold', color='#C00000')
ax.text(x[1] + width/2, electrico[1] + 50, 'Bs 600\n(76.0% Ahorro)', ha='center', fontweight='bold', color='#385723')

ax.set_ylim(0, 3000)
fig.tight_layout()
plt.savefig('grafico_tema2_comparativa_costos.png', dpi=300)
plt.close()
print("Gráfico 2 generado: grafico_tema2_comparativa_costos.png")


# 4. FLUJOGRAMA TEMA 2: MATRIZ Y RED DE ELECTROMOVILIDAD EN BOLIVIA
fig, ax = plt.subplots(figsize=(10, 3.5), dpi=300)
ax.axis('off')

boxes_ev = [
    ("1. Generación ENDE", "Centrales Hidroeléctricas\ny Parques Solares Nacionales"),
    ("2. Subestaciones Red 220V", "Distribución por Delapaz, CRE,\nElfec en Eje Troncal"),
    ("3. Electrineras / Toma 220V", "Estaciones de carga rápida DC\no toma doméstica de hogar"),
    ("4. Batería de Litio Quantum", "Almacenamiento eficiente LFP\ny movilidad cero emisiones")
]

colors_ev = ['#1F497D', '#2E75B6', '#00B0F0', '#385723']

for i, (title, desc) in enumerate(boxes_ev):
    rect = patches.FancyBboxPatch((i*2.4 + 0.1, 0.3), 2.0, 1.4, boxstyle="round,pad=0.1", ec="none", fc=colors_ev[i])
    ax.add_patch(rect)
    ax.text(i*2.4 + 1.1, 1.3, title, color='white', weight='bold', fontsize=10, ha='center', va='center')
    ax.text(i*2.4 + 1.1, 0.8, desc, color='white', fontsize=8, ha='center', va='center', multialignment='center')
    
    if i < len(boxes_ev) - 1:
        ax.annotate('', xy=(i*2.4 + 2.35, 1.0), xytext=(i*2.4 + 2.15, 1.0),
                    arrowprops=dict(arrowstyle="->", color='#333333', lw=2.5))

ax.set_xlim(0, 9.6)
ax.set_ylim(0, 2.0)
plt.title('Flujograma de Integración Energética y Electromovilidad en Bolivia (ENDE + Quantum)', fontsize=11, fontweight='bold', pad=10, color='#1F497D')
fig.tight_layout()
plt.savefig('grafico_tema2_flujograma.png', dpi=300)
plt.close()
print("Flujograma 2 generado: grafico_tema2_flujograma.png")
