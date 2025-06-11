import pandas as pd
from datetime import datetime, timedelta
import dash
from dash import dcc, html, Input, Output
import dash_bootstrap_components as dbc
import plotly.express as px

# Chargement des données
df = pd.read_excel('data/delegation0.xlsx', sheet_name='Sheet1')
print(df.columns.tolist())
# Nettoyage des noms de colonnes (élimine les espaces en trop)
df.columns = df.columns.str.strip()

# Conversion des dates
for col in ['date de mise en place', 'date d echeance', 'date de saisie',
            'date de derniere reevaluation', 'date de maturite de l engagement']:
    df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')

# Création d'une colonne de durée
df['duree_garantie'] = (df['date d echeance'] - df['date de mise en place']).dt.days

# Setup Dash + Bootstrap
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
server = app.server

# Dropdown options
type_opts = [{"label": t, "value": t} for t in df['libelle nature'].dropna().unique()]
assureur_opts = [{"label": t, "value": t} for t in df['nom garant'].dropna().unique()]
annees = df['date de mise en place'].dt.year.dropna().unique()
year_opts = [{"label": int(y), "value": int(y)} for y in sorted(annees)]

app.layout = dbc.Container(fluid=True, children=[
    dbc.Row(dbc.Col(html.H1("Dashboard Délégations d'Assurance Corporate", className="text-center my-4"))),

    dbc.Row([
        dbc.Col(dbc.Form([dbc.Label("Type de garantie"), dcc.Dropdown(type_opts, multi=True, id='filter-type')]), width=4),
        dbc.Col(dbc.Form([dbc.Label("Assureur"), dcc.Dropdown(assureur_opts, multi=True, id='filter-assureur')]), width=4),
        dbc.Col(dbc.Form([dbc.Label("Année mise en place"), dcc.Dropdown(year_opts, multi=True, id='filter-year')]), width=4),
    ], className="mb-4"),

    dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("Clients distincts"), dbc.CardBody(html.H4(id='total-clients'))]), md=2),
        dbc.Col(dbc.Card([dbc.CardHeader("Total garanties"), dbc.CardBody(html.H4(id='total-garanties'))]), md=3),
        dbc.Col(dbc.Card([dbc.CardHeader("Garanties actives"), dbc.CardBody(html.H4(id='garanties-actives'))]), md=3),
        dbc.Col(dbc.Card([dbc.CardHeader("Garanties expirées"), dbc.CardBody(html.H4(id='garanties-expires'))]), md=2),
        dbc.Col(dbc.Card([dbc.CardHeader("Garanties resiliées"), dbc.CardBody(html.H4(id='garanties-resilies'))]), md=2),
    ], className="mb-4"),

    dbc.Row([
        dbc.Col(dcc.Graph(id='repartition-type-garantie'), md=6),
        dbc.Col(dcc.Graph(id='repartition-segment'), md=6),
    ], className="mb-4"),

    dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("Montant total garanties"), dbc.CardBody(html.H5(id='montant-total-garanties'))]), md=3),
        dbc.Col(dbc.Card([dbc.CardHeader("Montant moyen"), dbc.CardBody(html.H5(id='montant-moyen-garantie'))]), md=3),
        dbc.Col(dbc.Card([dbc.CardHeader("Engagements couverts"), dbc.CardBody(html.H5(id='montant-engagement-couvert'))]), md=3),
        dbc.Col(dbc.Card([dbc.CardHeader("Ratio garantie / engagement"), dbc.CardBody(html.H5(id='ratio-garantie-engagement'))]), md=3),
    ], className="mb-4"),

    dbc.Row([
        dbc.Col(dcc.Graph(id='top10-garanties'), md=6),
        dbc.Col(dcc.Graph(id='repartition-montant-garanties'), md=6),
    ], className="mb-4"),

    dbc.Row([
        dbc.Col(dcc.Graph(id='historique-mises-en-place'), md=6),
        dbc.Col(dcc.Graph(id='echeances-par-mois'), md=6)
    ], className="mb-4"),

    dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("Échéance 3 📆"), dbc.CardBody(html.H5(id='echeance-3mois'))]), md=3),
        dbc.Col(dbc.Card([dbc.CardHeader("Renouvellements 30j"), dbc.CardBody(html.H5(id='garanties-renouvelees'))]), md=3),
        dbc.Col(dbc.Card([dbc.CardHeader("Durée moyenne"), dbc.CardBody(html.H5(id='duree-moyenne-garantie'))]), md=3),
        dbc.Col(dbc.Card([dbc.CardHeader("Délai saisie→mise en place"), dbc.CardBody(html.H5(id='delai-saisie-mise-en-place'))]), md=3),
    ], className="mb-4"),

    dbc.Row(dbc.Col(html.Footer("© 2025 Dashboard Assurance", className="text-center mt-4")))
])

@app.callback(
    Output('total-clients', 'children'),
    Output('total-garanties', 'children'),
    Output('garanties-actives', 'children'),
    Output('garanties-expires', 'children'),
    Output('garanties-resilies', 'children'),
    Output('repartition-type-garantie', 'figure'),
    Output('repartition-segment', 'figure'),
    Output('montant-total-garanties', 'children'),
    Output('montant-moyen-garantie', 'children'),
    Output('montant-engagement-couvert', 'children'),
    Output('ratio-garantie-engagement', 'children'),
    Output('top10-garanties', 'figure'),
    Output('repartition-montant-garanties', 'figure'),
    Output('historique-mises-en-place', 'figure'),
    Output('echeances-par-mois', 'figure'),
    Output('echeance-3mois', 'children'),
    Output('garanties-renouvelees', 'children'),
    Output('duree-moyenne-garantie', 'children'),
    Output('delai-saisie-mise-en-place', 'children'),
    Input('filter-type', 'value'),
    Input('filter-assureur', 'value'),
    Input('filter-year', 'value')
)
def update_all(type_filter, assur_filter, year_filter):
    dff = df.copy()
    if type_filter:
        dff = dff[dff['libelle nature'].isin(type_filter)]
    if assur_filter:
        dff = dff[dff['nom garant'].isin(assur_filter)]
    if year_filter:
        dff = dff[dff['date de mise en place'].dt.year.isin(year_filter)]

    today = datetime.today()

    clients = dff['code client'].nunique()
    total = len(dff)
    actives = len(dff[dff['situationde la garantie'] == 'Active'])
    expires = len(dff[dff['situationde la garantie'] =='Expirée'])
    resilies = len(dff[dff['situationde la garantie'] =='Résiliée'])

    fig_type = px.pie(dff, names='libelle nature', title='Répartition par type')
    fig_seg = px.pie(dff, names='libelle segment', title='Répartition par segment')

    mt_total_val = dff['montant de la garantie'].sum()
    mt_total = f"{mt_total_val:,.2f} €"
    mt_mean = f"{dff['montant de la garantie'].mean():,.2f} €"
    eng_val = dff['montant engagement couvert actualisé'].sum()
    mt_eng = f"{eng_val:,.2f} €"
    ratio = mt_total_val / eng_val if eng_val > 0 else 0
    ratio_str = f"{ratio:.2%}"

    fig_top10 = px.bar(
        dff.nlargest(10, 'montant de la garantie'),
        x='nom client', y='montant de la garantie', title="Top 10 garanties"
    )

    bins = [0, 1e6, 5e6, float('inf')]
    labels = ['<1M', '1-5M', '>5M']
    dff['tranche'] = pd.cut(dff['montant de la garantie'], bins=bins, labels=labels)
    fig_dist = px.pie(dff, names='tranche', title='Répartition montant')

    hist = dff.groupby(dff['date de mise en place'].dt.to_period('M')).size().reset_index(name='count')
    hist['date'] = hist['date de mise en place'].astype(str)
    fig_hist = px.line(hist, x='date', y='count', title="Historique mises en place")

    ech = dff.groupby(dff['date d echeance'].dt.to_period('M')).size().reset_index(name='count')
    ech['date'] = ech['date d echeance'].astype(str)
    fig_ech = px.bar(ech, x='date', y='count', title="Échéances par mois")

    echeance3 = len(dff[(dff['date d echeance'] >= today) & (dff['date d echeance'] <= today + timedelta(days=90))])
    renouv = len(dff[dff['date de mise en place'] >= today - timedelta(days=30)])
    duree = f"{dff['duree_garantie'].mean():.0f} j"
    delai = f"{(dff['date de mise en place'] - dff['date de saisie']).dt.days.mean():.0f} j"

    return (clients, total, actives, expires,resilies,
            fig_type, fig_seg,
            mt_total, mt_mean, mt_eng, ratio_str,
            fig_top10, fig_dist,
            fig_hist, fig_ech,
            echeance3, renouv, duree, delai)

if __name__ == "__main__":
    app.run(debug=True)
