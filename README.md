# Dashboard: Segmento EMPRESAS — Projeto Metaday 2025
**Recriação fiel do Power BI em Python (Streamlit + Plotly)**

---

## 🚀 Como executar

```bash
# 1. Instale as dependências
pip install -r requirements.txt

# 2. Execute o app
streamlit run app.py
```

Acesse em: http://localhost:8501

---

## 📁 Estrutura esperada do Excel

| Aba              | Descrição                                    |
|------------------|----------------------------------------------|
| `Empresas`       | Tabela principal (408 registros)             |
| `Ramo`           | CNAE, ID do ramo, sub-ramo e ramo principal  |
| `Sub-Ramos`      | ID + nome do sub-ramo                        |
| `Ramo Principal` | ID + nome do ramo principal                  |
| `Bairros`        | Bairro → Distrito                            |
| `Distrito`       | Dados IBGE/Censo 2022 por distrito           |
| `Porte`          | Tabela auxiliar de porte                     |
| `Horario`        | Tipos de expediente                          |

---

## 🗂️ Páginas do dashboard

| Página | Power BI         | Streamlit                                      |
|--------|------------------|------------------------------------------------|
| 1      | Visão Geral      | KPIs + Presença Digital + Porte + Mapa         |
| 2      | Distritos        | Treemap distritos + Tempo existência + Por ramo|
| 3      | FATEC            | Tempo/Distância FATEC + Sub-Ramos (Treemap)    |
| 4      | Mapa Ramos       | Mapa scatter por sub-ramo (filtro)             |
| 5      | Distritos/Segm.  | Treemap distritos filtrado por sub-ramo        |
| 6      | IBGE / Perfil    | Porte + Presença + Expediente + Tabela IBGE    |

---

## 🔄 Equivalências Power BI → Python

### Medidas DAX → Pandas

| DAX (Power BI)                              | Python / Pandas                                      |
|---------------------------------------------|------------------------------------------------------|
| `COUNTROWS(Empresas)`                       | `len(df)`                                            |
| `DISTINCTCOUNT(Empresas[Distrito])`         | `df["Distrito"].nunique()`                           |
| `CALCULATE(COUNT(...), filtro)`             | `df[df["col"] == valor].shape[0]`                    |
| `IF(ISBLANK([Site]), "NÃO", "SIM")`        | `df["Site"].apply(lambda x: "SIM" if ... else "NÃO")`|
| `DATEDIFF(Data Abertura, TODAY(), YEAR)`    | `(hoje - df["Data Abertura"]).dt.days / 365.25`      |
| Grupos de faixa de idade                    | Função `faixa_idade()` com `pd.cut` equivalente      |
| Grupos de distância                         | Função `faixa_dist()` com intervalos manuais         |
| Grupos de tempo                             | Função `faixa_tempo()` com intervalos manuais        |

### Visuais

| Power BI Visual      | Plotly / Streamlit                          |
|----------------------|---------------------------------------------|
| Cartão KPI           | HTML customizado com `st.markdown`          |
| Gráfico Pizza        | `px.pie()` com `hole=0.35` (donut)          |
| Barras horizontais   | `px.bar(orientation="h")`                   |
| Barras verticais     | `px.bar()`                                  |
| Treemap              | `px.treemap()` com `path` hierárquico       |
| Mapa (Azure Maps)    | `px.scatter_mapbox()` com OpenStreetMap     |
| Tabela               | `st.dataframe()`                            |
| Filtro/Slicer        | `st.multiselect()` na sidebar               |

---

## 💡 Sugestões de melhoria

1. **Geocodificação automática**: se as coordenadas não estiverem na planilha, use `geopy` + Nominatim para geocodificar por endereço.

2. **Mapa de calor**: substitua `scatter_mapbox` por `density_mapbox` para visualizar concentração de empresas.

3. **Filtro de data**: adicione um slider de `Data Abertura` para análise temporal.

4. **Exportação**: adicione botão `st.download_button` para exportar o df filtrado em CSV/Excel.

5. **Cache com TTL**: use `@st.cache_data(ttl=3600)` para atualizações automáticas se os dados forem dinâmicos.

6. **Deploy**: publique gratuitamente em [Streamlit Community Cloud](https://streamlit.io/cloud) com repositório GitHub.

7. **Tabela interativa**: substitua `st.dataframe` por `st.data_editor` para permitir edição inline.

8. **Alertas de oportunidade**: adicione lógica para destacar distritos com baixa presença digital (< 50%) como oportunidades de mercado.
