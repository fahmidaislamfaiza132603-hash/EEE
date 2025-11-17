import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io
import snowflake.connector  # For Snowflake; pip install snowflake-connector-python
from datetime import datetime

# Snowflake connection (replace with your credentials)
def get_snowflake_conn():
    return snowflake.connector.connect(
        user='YOUR_USER',
        password='YOUR_PASSWORD',
        account='YOUR_ACCOUNT_URL',  # e.g., 'youraccount.snowflakecomputing.com'
        warehouse='YOUR_WAREHOUSE',
        database='SPORTS_DB'
    )

# Load data from Snowflake
def load_teams():
    conn = get_snowflake_conn()
    df = pd.read_sql("SELECT * FROM TEAMS", conn)
    conn.close()
    return df.to_dict('records')

def load_matches():
    conn = get_snowflake_conn()
    df = pd.read_sql("SELECT * FROM MATCHES", conn)
    conn.close()
    return df.to_dict('records')

def load_schedule():
    conn = get_snowflake_conn()
    df = pd.read_sql("SELECT * FROM SCHEDULE", conn)
    conn.close()
    return df.to_dict('records')

# Save data to Snowflake
def save_teams(teams):
    conn = get_snowflake_conn()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM TEAMS")  # Clear and re-insert (simple; optimize for production)
    for t in teams:
        cursor.execute("INSERT INTO TEAMS VALUES (%s, %s, %s, %s, %s, %s, %s, %s)",
                       (t['name'], t['captain'], t['contact'], t['players'], t['matches'], t['wins'], t['losses'], t['points']))
    conn.commit()
    conn.close()

def save_matches(matches):
    conn = get_snowflake_conn()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM MATCHES")
    for m in matches:
        cursor.execute("INSERT INTO MATCHES VALUES (%s, %s, %s, %s, %s, %s, %s)",
                       (m['game'], m['team1'], m['team2'], m['date'], m['winner'], m['score'], m['venue']))
    conn.commit()
    conn.close()

def save_schedule(schedules):
    conn = get_snowflake_conn()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM SCHEDULE")
    for s in schedules:
        cursor.execute("INSERT INTO SCHEDULE VALUES (%s, %s, %s, %s)",
                       (s['date'], s['team1'], s['team2'], s['venue']))
    conn.commit()
    conn.close()

# If no Snowflake, use this instead (comment out Snowflake functions and use these)
# teams_data = []
# matches_data = []
# schedule_data = []

# Authentication
def login():
    st.title("Stamford University EEE Sports Management")
    username = st.text_input("User ID")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if username == "admin" and password == "teacher":
            st.session_state['role'] = 'admin'
            st.success("Logged in as Admin/Teacher")
        elif username == "visitor" and password == "view":
            st.session_state['role'] = 'visitor'
            st.success("Logged in as Visitor")
        else:
            st.error("Invalid credentials")

# Home Screen
def home():
    st.title("EEE Department Sports Dashboard")
    teams = load_teams()
    matches = load_matches()
    schedules = load_schedule()
    active = len([s for s in schedules if s['date'] >= str(datetime.now().date())])
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Teams", len(teams))
    col2.metric("Total Matches", len(matches))
    col3.metric("Active Games", active)

# Add Match
def add_match():
    st.title("Add Match Result")
    game = st.selectbox("Game", ["Cricket", "Football", "Badminton", "Chess", "Carrom", "Ludo", "Table Tennis", "Basketball", "Volleyball", "Other"])
    team1 = st.text_input("Team 1")
    team2 = st.text_input("Team 2")
    date = st.date_input("Date")
    winner = st.selectbox("Winner", ["Team 1", "Team 2", "Draw"])
    score = st.text_input("Score")
    venue = st.text_input("Venue")
    if st.button("Save Match"):
        match = {"game": game, "team1": team1, "team2": team2, "date": str(date), "winner": winner, "score": score, "venue": venue}
        matches = load_matches()
        matches.append(match)
        save_matches(matches)
        # Update teams
        teams = load_teams()
        for t in teams:
            if t['name'] in [team1, team2]:
                t['matches'] += 1
                if winner == "Team 1" and t['name'] == team1:
                    t['wins'] += 1
                    t['points'] += 2
                elif winner == "Team 2" and t['name'] == team2:
                    t['wins'] += 1
                    t['points'] += 2
                elif winner == "Draw":
                    t['points'] += 1
                else:
                    t['losses'] += 1
        save_teams(teams)
        st.success("Match saved")

# Scoreboard
def scoreboard():
    st.title("Live Scoreboard")
    game_filter = st.text_input("Filter by Game")
    teams = load_teams()
    if game_filter:
        teams = [t for t in teams if game_filter.lower() in t.get('game', '').lower()]
    df = pd.DataFrame(teams)
    st.dataframe(df[['name', 'matches', 'wins', 'losses', 'points']])

# Schedule
def schedule():
    st.title("Match Schedule")
    schedules = load_schedule()
    df = pd.DataFrame(schedules)
    st.dataframe(df)
    if st.session_state.get('role') == 'admin':
        st.subheader("Add New Match")
        date = st.date_input("Date")
        team1 = st.text_input("Team 1")
        team2 = st.text_input("Team 2")
        venue = st.text_input("Venue")
        if st.button("Add Schedule"):
            sched = {"date": str(date), "team1": team1, "team2": team2, "venue": venue}
            schedules.append(sched)
            save_schedule(schedules)
            st.success("Schedule added")

# Team Management
def team_mgmt():
    st.title("Team Management")
    teams = load_teams()
    df = pd.DataFrame(teams)
    st.dataframe(df)
    st.subheader("Add Team")
    name = st.text_input("Team Name")
    captain = st.text_input("Captain")
    contact = st.text_input("Contact")
    players = st.number_input("Players", min_value=1)
    if st.button("Add Team"):
        team = {"name": name, "captain": captain, "contact": contact, "players": players, "matches": 0, "wins": 0, "losses": 0, "points": 0}
        teams.append(team)
        save_teams(teams)
        st.success("Team added")

# Export
def export():
    st.title("Export & Print")
    matches = load_matches()
    df = pd.DataFrame(matches)
    csv = df.to_csv(index=False)
    st.download_button("Download CSV", csv, "matches.csv", "text/csv")
    # PDF
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    c.drawString(100, 750, "Scoreboard")
    teams = load_teams()
    y = 700
    for t in teams:
        c.drawString(100, y, f"{t['name']}: {t['points']} points")
        y -= 20
    c.save()
    buffer.seek(0)
    st.download_button("Download PDF", buffer, "scoreboard.pdf", "application/pdf")

# Stats
def stats():
    st.title("Statistics")
    matches = load_matches()
    game_counts = {}
    for m in matches:
        game_counts[m['game']] = game_counts.get(m['game'], 0) + 1
    most_game = max(game_counts, key=game_counts.get) if game_counts else "None"
    teams = load_teams()
    top_team = max(teams, key=lambda t: t['points'])['name'] if teams else "None"
    st.write(f"Most Played Game: {most_game}")
    st.write(f"Most Successful Team: {top_team}")

# Main App
def main():
    st.set_page_config(page_title="EEE Sports", page_icon="🏆", layout="wide")
    if 'role' not in st.session_state:
        login()
    else:
        menu = ["Home", "Scoreboard", "Stats"]
        if st.session_state['role'] == 'admin':
            menu += ["Add Match", "Schedule", "Team Mgmt", "Export"]
        choice = st.sidebar.selectbox("Menu", menu)
        if choice == "Home":
            home()
        elif choice == "Add Match":
            add_match()
        elif choice == "Scoreboard":
            scoreboard()
        elif choice == "Schedule":
            schedule()
        elif choice == "Team Mgmt":
            team_mgmt()
        elif choice == "Export":
            export()
        elif choice == "Stats":
            stats()
        if st.sidebar.button("Logout"):
            st.session_state.clear()

if __name__ == "__main__":
    main()
