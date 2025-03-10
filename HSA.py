import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image
import os

# Configuración de la página
st.set_page_config(
    page_title="Dashboard Expedientes H.M. CESAR LORDUY",
    page_icon="📊",
    layout="wide"
)

# Título y descripción
st.title("📊 Dashboard de Expedientes")
st.markdown("### Análisis de expedientes del H.M. CESAR LORDUY")

# Botón para limpiar caché y forzar recarga
if st.sidebar.button("🔄 Recargar datos", use_container_width=True):
    # Limpiar el caché
    st.cache_data.clear()
    # Reiniciar la aplicación
    st.rerun()

# Carga de datos
@st.cache_data
def load_data():
    # Definir la ruta al archivo Excel
    file_path = r"C:\Users\mateo\Downloads\HSA EXPEDEDIENTES VIVOS\HSA PRUEBA 26.xlsx"
    
    # Cargar el archivo Excel
    df = pd.read_excel(file_path, sheet_name="HSA")
    return df

# Cargar los datos
try:
    df = load_data()
    
    # Mostrar mensaje de éxito
    st.success("Datos cargados correctamente.")
    
    # Crear un contenedor para organizar los mosaicos
    col1, col2, col3 = st.columns(3)
    
    # Total de expedientes
    total_expedientes = len(df)
    
    # Contar expedientes por estado
    estado_counts = df['ESTADO'].value_counts()
    despacho_count = estado_counts.get('DESPACHO', 0)
    prearchivo_count = estado_counts.get('PRE-ARCHIVO', 0)
    
    # Calcular porcentajes
    porcentaje_despacho = (despacho_count / total_expedientes) * 100
    porcentaje_prearchivo = (prearchivo_count / total_expedientes) * 100
    
    # Contar asesores únicos
    total_asesores = df['ASESOR'].nunique()
    
    # Crear mosaicos con contadores
    with col1:
        st.metric(
            label="Total Expedientes",
            value=f"{total_expedientes}",
            delta=None
        )
        
    with col2:
        st.metric(
            label="Expedientes en DESPACHO",
            value=f"{despacho_count}",
            delta=f"{porcentaje_despacho:.1f}% del total"
        )
        
    with col3:
        st.metric(
            label="Expedientes en PRE-ARCHIVO",
            value=f"{prearchivo_count}",
            delta=f"{porcentaje_prearchivo:.1f}% del total"
        )
    
    # Crear nueva fila para más información
    col1, col2 = st.columns(2)
    
    # Gráfico de torta para DESPACHO y PRE-ARCHIVO
    with col1:
        st.subheader("Distribución de Expedientes")
        
        # Datos para el gráfico
        labels = ['DESPACHO', 'PRE-ARCHIVO']
        values = [despacho_count, prearchivo_count]
        
        # Crear figura con plotly
        fig = px.pie(
            names=labels,
            values=values,
            title="Porcentaje de Expedientes por Estado",
            color_discrete_sequence=px.colors.qualitative.Bold,
            hole=0.4
        )
        
        # Añadir anotaciones
        fig.update_traces(
            textposition='inside', 
            textinfo='percent+label',
            marker=dict(line=dict(color='#FFFFFF', width=2))
        )
        
        fig.update_layout(
            legend_title="Estado",
            font=dict(size=14)
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    # Tabla con conteo de expedientes por asesor
    with col2:
        st.subheader(f"Asesores ({total_asesores})")
        
        # Conteo de expedientes por asesor y estado
        asesor_counts = df.groupby('ASESOR')['EXPEDIENTE'].count().reset_index()
        asesor_counts.columns = ['Asesor', 'Total Expedientes']
        
        # Obtener conteo por estado para cada asesor
        despacho_por_asesor = df[df['ESTADO'] == 'DESPACHO'].groupby('ASESOR')['EXPEDIENTE'].count().reset_index()
        despacho_por_asesor.columns = ['Asesor', 'DESPACHO']
        
        prearchivo_por_asesor = df[df['ESTADO'] == 'PRE-ARCHIVO'].groupby('ASESOR')['EXPEDIENTE'].count().reset_index()
        prearchivo_por_asesor.columns = ['Asesor', 'PRE-ARCHIVO']
        
        # Unir las tablas
        asesor_tabla = pd.merge(asesor_counts, despacho_por_asesor, on='Asesor', how='left')
        asesor_tabla = pd.merge(asesor_tabla, prearchivo_por_asesor, on='Asesor', how='left')
        
        # Llenar NaN con 0
        asesor_tabla = asesor_tabla.fillna(0)
        
        # Convertir a enteros
        asesor_tabla['DESPACHO'] = asesor_tabla['DESPACHO'].astype(int)
        asesor_tabla['PRE-ARCHIVO'] = asesor_tabla['PRE-ARCHIVO'].astype(int)
        
        # Ordenar por total de expedientes
        asesor_tabla = asesor_tabla.sort_values('Total Expedientes', ascending=False)
        
        # Mostrar tabla
        st.dataframe(asesor_tabla, use_container_width=True)
    
    # Visualización adicional - Gráfico de barras para expedientes por asesor
    st.subheader("Distribución de Expedientes por Asesor")
    
    # Preparar datos para el gráfico
    despacho_df = df[df['ESTADO'] == 'DESPACHO'].groupby('ASESOR').size().reset_index(name='DESPACHO')
    prearchivo_df = df[df['ESTADO'] == 'PRE-ARCHIVO'].groupby('ASESOR').size().reset_index(name='PRE-ARCHIVO')
    
    # Combinar ambos dataframes
    merged_df = pd.merge(despacho_df, prearchivo_df, on='ASESOR', how='outer').fillna(0)
    
    # Crear gráfico de barras apiladas
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=merged_df['ASESOR'],
        y=merged_df['DESPACHO'],
        name='DESPACHO',
        marker_color='#1f77b4'
    ))
    
    fig.add_trace(go.Bar(
        x=merged_df['ASESOR'],
        y=merged_df['PRE-ARCHIVO'],
        name='PRE-ARCHIVO',
        marker_color='#ff7f0e'
    ))
    
    fig.update_layout(
        barmode='stack',
        title='Expedientes por Asesor y Estado',
        xaxis={'title': 'Asesor', 'categoryorder': 'total descending'},
        yaxis={'title': 'Número de Expedientes'},
        legend_title='Estado',
        hovermode='closest'
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # NUEVO: Gráfico de barras por tema
    st.subheader("Distribución de Expedientes por Tema")
    
    # Verificar si existe la columna TEMA
    if 'TEMA' in df.columns:
        # Contar expedientes por tema y estado
        tema_df = df.groupby(['TEMA', 'ESTADO']).size().reset_index(name='CANTIDAD')
        
        # Filtrar temas sin valores nulos o vacíos
        tema_df = tema_df[tema_df['TEMA'].notna() & (tema_df['TEMA'] != '')]
        
        if not tema_df.empty:
            # Crear gráfico de barras agrupadas
            fig_tema = px.bar(
                tema_df,
                x='TEMA',
                y='CANTIDAD',
                color='ESTADO',
                title='Distribución de Expedientes por Tema y Estado',
                barmode='group',
                color_discrete_map={
                    'DESPACHO': '#1f77b4',
                    'PRE-ARCHIVO': '#ff7f0e'
                },
                text_auto=True
            )
            
            fig_tema.update_layout(
                xaxis_title='Tema',
                yaxis_title='Número de Expedientes',
                legend_title='Estado',
                xaxis={'categoryorder': 'total descending'},
                font=dict(size=12)
            )
            
            # Mejoras visuales
            fig_tema.update_traces(
                textposition='outside',
                textfont_size=10
            )
            
            # Mostrar gráfico
            st.plotly_chart(fig_tema, use_container_width=True)
            
            # Mostrar tabla resumen por tema
            with st.expander("Ver resumen por tema"):
                # Crear tabla resumen
                tema_total = df.groupby('TEMA')['EXPEDIENTE'].count().reset_index()
                tema_total.columns = ['Tema', 'Total Expedientes']
                tema_total = tema_total.sort_values('Total Expedientes', ascending=False)
                
                # Mostrar tabla
                st.dataframe(tema_total, use_container_width=True)
        else:
            st.warning("No se encontraron datos de temas válidos para generar el gráfico.")
    else:
        st.warning("No se encontró la columna 'TEMA' en los datos.")
    
    # Mostrar información adicional
    with st.expander("Ver datos completos"):
        st.dataframe(df)
        
except Exception as e:
    st.error(f"Error al cargar o procesar los datos: {e}")
    st.info("Por favor, asegúrate de que el archivo 'HSA H.M. CESAR LORDUY 2025.xlsx' esté presente en la ruta: C:\\Users\\mateo\\Downloads\\HSA EXPEDEDIENTES VIVOS")