import dash
from dash import dcc, html
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px

# Chargement des données
df = pd.read_excel('data/delegation0.xlsx', sheet_name='Sheet1')
df.columns = df.columns.str.strip()

# Conversion des dates
for col in ['date de mise en place', 'date d echeance', 'date de saisie',
            'date de derniere reevaluation', 'date de maturite de l engagement']:
    df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')

df['duree_garantie'] = (df['date d echeance'] - df['date de mise en place']).dt.days

# Setup Dash
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])#QUARTZ  SKETCHY MINTY SOLAR
server = app.server


from datetime import datetime, timedelta
import dash
from dash import dash_table

from datetime import datetime, timedelta
import pandas as pd
import dash
from dash import dash_table

def create_echeance_table(dff):
    today = datetime.today()

    # Copie sécurisée et conversion des dates
    dff = dff.copy()
    dff['date d echeance'] = pd.to_datetime(dff['date d echeance'], errors='coerce')
    dff['date de mise en place'] = pd.to_datetime(dff['date de mise en place'], errors='coerce')

    # Filtrage des données dans les 3 mois
    mask = (dff['date d echeance'] >= today) & (dff['date d echeance'] <= today + timedelta(days=90))
    echeance_3mois = dff.loc[mask].copy()

    if echeance_3mois.empty:
        return dash_table.DataTable(
            data=[],
            columns=[{'name': 'Aucune échéance à venir', 'id': 'info'}],
            style_cell={'textAlign': 'center', 'fontStyle': 'italic'}
        )

    # Calcul des jours restants
    echeance_3mois['Jours restants'] = (echeance_3mois['date d echeance'] - today).dt.days
    echeance_3mois = echeance_3mois.sort_values('Jours restants')

    # Formatage des colonnes
    echeance_3mois['date de mise en place'] = echeance_3mois['date de mise en place'].dt.strftime('%d/%m/%Y')
    echeance_3mois['date d echeance'] = echeance_3mois['date d echeance'].dt.strftime('%d/%m/%Y')

    if 'montant de la garantie' in echeance_3mois.columns:
        echeance_3mois['montant de la garantie'] = echeance_3mois['montant de la garantie'].apply(
            lambda x: f"{x:,.2f} €" if pd.notnull(x) else ""
        )

    # Colonnes à afficher et leur mapping
    cols_to_display = {
        'nom client': 'Client',
        'libelle nature': 'Type garantie',
        'montant de la garantie': 'Montant',
        'date de mise en place': 'Mise en place',
        'date d echeance': 'Échéance',
        'Jours restants': 'Jours restants',
        'nom garant': 'Assureur'
    }

    df_final = echeance_3mois.rename(columns=cols_to_display)

    return dash_table.DataTable(
        data=df_final.to_dict('records'),
        columns=[{'name': v, 'id': v} for v in cols_to_display.values()],
        
        style_table={
            'overflowX': 'auto',
            'borderRadius': '10px',
            'boxShadow': '0 2px 4px rgba(0,0,0,0.1)',
            'fontFamily': 'Arial, sans-serif',
            'margin': '20px 0'
        },
        
        style_header={
            'backgroundColor': '#3498db',
            'color': 'white',
            'fontWeight': 'bold',
            'textAlign': 'center',
            'border': 'none',
            'padding': '12px 15px'
        },
        
        style_cell={
            'padding': '10px 15px',
            'textAlign': 'left',
            'borderBottom': '1px solid #ecf0f1',
            'minWidth': '120px',
            'maxWidth': '200px',
            'whiteSpace': 'normal',
            'fontSize': '14px'
        },
        
        style_data={
            'backgroundColor': 'white',
            'color': '#34495e',
            'transition': 'background-color 0.3s ease'
        },
        
        style_data_conditional=[
            {
                'if': {
                    'filter_query': '{Jours restants} < 7',
                    'column_id': 'Jours restants'
                },
                'backgroundColor': '#ffebee',
                'color': '#c62828',
                'fontWeight': 'bold',
                'borderLeft': '3px solid #c62828'
            },
            {
                'if': {
                    'filter_query': '{Jours restants} >= 7 && {Jours restants} < 15',
                    'column_id': 'Jours restants'
                },
                'backgroundColor': '#fff8e1',
                'color': '#ff8f00',
                'borderLeft': '3px solid #ff8f00'
            },
            {
                'if': {'row_index': 'odd'},
                'backgroundColor': '#f8f9fa'
            },
            {
                'if': {'state': 'selected'},
                'backgroundColor': '#bde0fe',
                'border': '1px solid #3498db'
            },
            {
                'if': {'column_id': ['Montant', 'Jours restants']},
                'textAlign': 'center'
            },
            {
                'if': {'column_id': 'Montant'},
                'fontWeight': 'bold',
                'color': '#2e7d32'
            }
        ],
        
        page_size=10,
        filter_action='native',
        sort_action='native',
        sort_mode='multi',
        page_action='native',
        
        tooltip_data=[
            {
                column: {'value': str(value), 'type': 'markdown'}
                for column, value in row.items()
            } for row in df_final.to_dict('records')
        ],
        tooltip_duration=None,
        
        style_cell_conditional=[
            {
                'if': {'column_id': 'Échéance'},
                'fontStyle': 'italic'
            },
            {
                'if': {'column_id': 'Client'},
                'fontWeight': '600'
            }
        ]
    )



def render_general_tab(dff):
    today = datetime.today()
    
    # Calcul des indicateurs principaux
    clients = dff['code client'].nunique()
    total = len(dff)
    actives = len(dff[dff['situationde la garantie'] == 'Active'])
    expires = len(dff[dff['situationde la garantie'] == 'Expirée'])
    resilies = len(dff[dff['situationde la garantie'] == 'Résiliée'])
    
    # Calcul des montants
    mt_total_val = dff['montant de la garantie'].sum()
    mt_total = f"{mt_total_val:,.2f} €"
    mt_mean = f"{dff['montant de la garantie'].mean():,.2f} €"
    eng_val = dff['montant engagement couvert actualisé'].sum()
    mt_eng = f"{eng_val:,.2f} €"
    ratio = mt_total_val / eng_val if eng_val > 0 else 0
    ratio_str = f"{ratio:.2%}"
    
    # Calcul des indicateurs temporels
    echeance3 = len(dff[(dff['date d echeance'] >= today) & (dff['date d echeance'] <= today + timedelta(days=90))])
    renouv = len(dff[dff['date de mise en place'] >= today - timedelta(days=30)])
    duree = f"{dff['duree_garantie'].mean():.0f} j"
    delai = f"{(dff['date de mise en place'] - dff['date de saisie']).dt.days.mean():.0f} j"
    
    # Création des graphiques
    fig_type = px.pie(dff, names='libelle nature', title='Répartition par type de garantie')
    fig_seg = px.pie(dff, names='libelle segment', title='Répartition par segment')
    
    fig_top10 = px.bar(
        dff.nlargest(10, 'montant de la garantie'),
        x='nom client', y='montant de la garantie', 
        title="Top 10 garanties par montant",
        height=400
    )
    
    bins = [0, 1e6, 5e6, float('inf')]
    labels = ['<1M', '1-5M', '>5M']
    dff['tranche'] = pd.cut(dff['montant de la garantie'], bins=bins, labels=labels)
    fig_dist = px.pie(dff, names='tranche', title='Répartition par montant')
    
    hist = dff.groupby(dff['date de mise en place'].dt.to_period('M')).size().reset_index(name='count')
    hist['date'] = hist['date de mise en place'].astype(str)
    fig_hist = px.line(hist, x='date', y='count', title="Historique des mises en place")
    
    ech = dff.groupby(dff['date d echeance'].dt.to_period('M')).size().reset_index(name='count')
    ech['date'] = ech['date d echeance'].astype(str)
    fig_ech = px.bar(ech, x='date', y='count', title="Échéances par mois")
    
    # Création du tableau des échéances
    echeance_table = create_echeance_table(dff)
    
    return html.Div([
        dbc.Row([
            dbc.Col(dbc.Card([dbc.CardHeader("Clients distincts"), dbc.CardBody(html.H4(clients))]), md=2),
            dbc.Col(dbc.Card([dbc.CardHeader("Total garanties"), dbc.CardBody(html.H4(total))]), md=3),
            dbc.Col(dbc.Card([dbc.CardHeader("Garanties actives"), dbc.CardBody(html.H4(actives))]), md=3),
            dbc.Col(dbc.Card([dbc.CardHeader("Garanties expirées"), dbc.CardBody(html.H4(expires))]), md=2),
            dbc.Col(dbc.Card([dbc.CardHeader("Garanties resiliées"), dbc.CardBody(html.H4(resilies))]), md=2),
        ], className="mb-4"),

        dbc.Row([
            dbc.Col(dcc.Graph(figure=fig_type), md=6),
            dbc.Col(dcc.Graph(figure=fig_seg), md=6),
        ], className="mb-4"),

        dbc.Row([
            dbc.Col(dbc.Card([dbc.CardHeader("Montant total garanties"), dbc.CardBody(html.H5(mt_total))]), md=3),
            dbc.Col(dbc.Card([dbc.CardHeader("Montant moyen"), dbc.CardBody(html.H5(mt_mean))]), md=3),
            dbc.Col(dbc.Card([dbc.CardHeader("Engagements couverts"), dbc.CardBody(html.H5(mt_eng))]), md=3),
            dbc.Col(dbc.Card([dbc.CardHeader("Ratio garantie / engagement"), dbc.CardBody(html.H5(ratio_str))]), md=3),
        ], className="mb-4"),

        dbc.Row([
            dbc.Col(dcc.Graph(figure=fig_top10), md=6),
            dbc.Col(dcc.Graph(figure=fig_dist), md=6),
        ], className="mb-4"),

        dbc.Row([
            dbc.Col(dcc.Graph(figure=fig_hist), md=6),
            dbc.Col(dcc.Graph(figure=fig_ech), md=6)
        ], className="mb-4"),

        dbc.Row([
            dbc.Col(dbc.Card([dbc.CardHeader("Échéance 3 mois"), dbc.CardBody(html.H5(echeance3))]), md=3),
            dbc.Col(dbc.Card([dbc.CardHeader("Renouvellements 30j"), dbc.CardBody(html.H5(renouv))]), md=3),
            dbc.Col(dbc.Card([dbc.CardHeader("Durée moyenne"), dbc.CardBody(html.H5(duree))]), md=3),
            dbc.Col(dbc.Card([dbc.CardHeader("Délai saisie→mise en place"), dbc.CardBody(html.H5(delai))]), md=3),
        ], className="mb-4"),
        
        # Nouvelle section pour le tableau des échéances
        dbc.Row([
            dbc.Col(html.H4("Détail des échéances dans 3 mois", className="mt-4"), width=12),
            dbc.Col(echeance_table, width=12)
        ], className="mb-4"),

        dbc.Row(dbc.Col(html.Footer("© 2025 Dashboard Assurance", className="text-center mt-4")))
    ])


def render_financier_tab(dff):
    mt_total_val = dff['montant de la garantie'].sum()
    mt_total = f"{mt_total_val:,.2f} €"
    mt_mean = f"{dff['montant de la garantie'].mean():,.2f} €"
    eng_val = dff['montant engagement couvert actualisé'].sum()
    mt_eng = f"{eng_val:,.2f} €"
    ratio = mt_total_val / eng_val if eng_val > 0 else 0
    ratio_str = f"{ratio:.2%}"

    return dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("Montant total garanties"), dbc.CardBody(html.H4(mt_total))]), md=3),
        dbc.Col(dbc.Card([dbc.CardHeader("Montant moyen par garantie"), dbc.CardBody(html.H4(mt_mean))]), md=3),
        dbc.Col(dbc.Card([dbc.CardHeader("Montant engagements couverts"), dbc.CardBody(html.H4(mt_eng))]), md=3),
        dbc.Col(dbc.Card([dbc.CardHeader("Ratio garantie / engagement"), dbc.CardBody(html.H4(ratio_str))]), md=3),
    ])

def render_partenaire_tab(dff):
    # Montant total garanti par assureur
    mt_par_assureur = dff.groupby('nom garant')['montant de la garantie'].sum().reset_index()
    nbr_garantie = dff.groupby('nom garant').size().reset_index(name='nbr_garantie')

    # Taux de renouvellement par assureur (nombre garanties mises en place dernièrement / total garanties)
    today = datetime.today()
    dernier_30j = dff[dff['date de mise en place'] >= today - timedelta(days=30)]
    renouv_par_assureur = dernier_30j.groupby('nom garant').size().reset_index(name='renouvellements')
    total_par_assureur = dff.groupby('nom garant').size().reset_index(name='total_garanties')

    taux_renouv = renouv_par_assureur.merge(total_par_assureur, on='nom garant', how='left')
    taux_renouv['tx_renouvellement'] = taux_renouv['renouvellements'] / taux_renouv['total_garanties']

    fig_mt = px.bar(mt_par_assureur, x='nom garant', y='montant de la garantie', title="Montant total garanti par assureur")
    fig_nbr = px.bar(nbr_garantie, x='nom garant', y='nbr_garantie', title="Nombre de garanties par assureur")
    fig_tx = px.bar(taux_renouv, x='nom garant', y='tx_renouvellement', title="Taux de renouvellement par assureur")

    return dbc.Row([
        dbc.Col(dcc.Graph(figure=fig_mt), md=4),
        dbc.Col(dcc.Graph(figure=fig_nbr), md=4),
        dbc.Col(dcc.Graph(figure=fig_tx), md=4),
    ])

def render_reevaluation_tab(dff):
    today = datetime.today()
    reevalues_30j = dff[dff['date de derniere reevaluation'] >= today - timedelta(days=30)]
    nb_reevalues = len(reevalues_30j)
    mt_reevalues = reevalues_30j['montant de la garantie'].sum()

    fig_nb = px.histogram(reevalues_30j, x='nom client', title="Nombre de réévaluations (30 derniers jours)", 
                          labels={'nom client': 'Client'}, nbins=20)
    
    fig_mt = px.histogram(reevalues_30j, x='nom client', y='montant de la garantie', 
                          title="Montant des garanties réévaluées", 
                          labels={'montant de la garantie': 'Montant (€)', 'nom client': 'Client'}, 
                          histfunc='sum')

    return html.Div([
        html.H4("Réévaluations récentes", className="mt-4"),
        dbc.Row([
            dbc.Col(dbc.Card([
                dbc.CardHeader("Nombre de garanties réévaluées (30j)"),
                dbc.CardBody(html.H4(nb_reevalues))
            ]), md=6),
            dbc.Col(dbc.Card([
                dbc.CardHeader("Montant total réévalué"),
                dbc.CardBody(html.H4(f"{mt_reevalues:,.2f} €"))
            ]), md=6)
        ], className="mb-4"),
        dbc.Row([
            dbc.Col(dcc.Graph(figure=fig_nb), md=6),
            dbc.Col(dcc.Graph(figure=fig_mt), md=6),
        ])
    ])


# Callback d'affichage des tabs avec filtres
@app.callback(
    Output('tab-content', 'children'),
    Input('tabs', 'active_tab'),
    Input('filter-type', 'value'),
    Input('filter-assureur', 'value'),
    Input('filter-year', 'value'),
)
def render_tab_content(active_tab, type_filter, assur_filter, year_filter):
    dff = df.copy()
    if type_filter:
        dff = dff[dff['libelle nature'].isin(type_filter)]
    if assur_filter:
        dff = dff[dff['nom garant'].isin(assur_filter)]
    if year_filter:
        dff = dff[dff['date de mise en place'].dt.year.isin(year_filter)]

    if active_tab == "tab-general":
        return render_general_tab(dff)
    elif active_tab == "tab-financier":
        return render_financier_tab(dff)
    elif active_tab == "tab-partenaire":
        return render_partenaire_tab(dff)
    elif active_tab == "tab-reevaluation":
        return render_reevaluation_tab(dff)
    return html.Div("Sélectionnez un onglet.")

# Layout principal
app.layout = dbc.Container([
    html.H1("Tableau de Bord - xxx", className="text-center my-4"),
    dbc.Row([
        dbc.Col([
            dcc.Dropdown(id='filter-type', options=[{'label': i, 'value': i} for i in df['libelle nature'].dropna().unique()],multi=True, placeholder="Filtrer par type de garantie"), ], md=4),
        dbc.Col([
            dcc.Dropdown(id='filter-assureur', options=[{'label': i, 'value': i} for i in df['nom garant'].dropna().unique()],
                         multi=True, placeholder="Filtrer par assureur"),
        ], md=4),
        dbc.Col([
            dcc.Dropdown(id='filter-year', options=[{'label': str(i), 'value': i} for i in sorted(df['date de mise en place'].dt.year.dropna().unique())],
                         multi=True, placeholder="Filtrer par année"),
        ], md=4),
    ], className="mb-4"),

    dbc.Tabs(id='tabs', active_tab='tab-general', children=[
        dbc.Tab(label="Vue Générale", tab_id='tab-general'),
        dbc.Tab(label="Vue Financière", tab_id='tab-financier'),
        dbc.Tab(label="Vue Assurance", tab_id='tab-partenaire'),
        dbc.Tab(label="Vue Réevaluation", tab_id='tab-reevaluation'),
    ]),

    html.Div(id='tab-content', className="mt-4"),
])

if __name__ == "__main__":
    app.run(debug=True)