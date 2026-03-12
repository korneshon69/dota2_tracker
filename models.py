from datetime import datetime
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()


class User(UserMixin, db.Model):
    __tablename__ = 'users'

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(64), unique=True, nullable=False, index=True)
    email = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(256), nullable=False)
    is_admin = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    is_active_user = db.Column(db.Boolean, default=True)

    matches = db.relationship('Match', backref='player', lazy='dynamic',
                              cascade='all, delete-orphan')

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def get_winrate(self):
        total = self.matches.count()
        if total == 0:
            return 0
        wins = self.matches.filter_by(result='win').count()
        return round(wins / total * 100, 1)

    def get_avg_kda(self):
        matches = self.matches.all()
        if not matches:
            return 0
        total_kda = sum(
            (m.kills + m.assists) / max(m.deaths, 1) for m in matches
        )
        return round(total_kda / len(matches), 2)

    def get_best_hero(self):
        results = db.session.query(
            Hero.name,
            db.func.count(Match.id).label('total'),
            db.func.sum(db.case((Match.result == 'win', 1), else_=0)).label('wins')
        ).join(Match, Match.hero_id == Hero.id).filter(
            Match.user_id == self.id
        ).group_by(Hero.id).having(
            db.func.count(Match.id) >= 1
        ).all()

        if not results:
            return None

        best = max(results, key=lambda r: r.wins / r.total if r.total > 0 else 0)
        winrate = round(best.wins / best.total * 100, 1) if best.total > 0 else 0
        return {'name': best.name, 'winrate': winrate, 'games': best.total}

    def __repr__(self):
        return f'<User {self.username}>'


class HeroRole(db.Model):
    __tablename__ = 'hero_roles'

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False)

    heroes = db.relationship('Hero', backref='role_ref', lazy='dynamic')

    def __repr__(self):
        return f'<HeroRole {self.name}>'


class MatchType(db.Model):
    __tablename__ = 'match_types'

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False)

    matches = db.relationship('Match', backref='match_type_ref', lazy='dynamic')

    def __repr__(self):
        return f'<MatchType {self.name}>'


class Hero(db.Model):
    __tablename__ = 'heroes'

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False, index=True)
    attribute = db.Column(db.String(20), nullable=False)  # strength, agility, intelligence, universal
    role_id = db.Column(db.Integer, db.ForeignKey('hero_roles.id'), nullable=False)
    difficulty = db.Column(db.Integer, nullable=False)  # 1-3
    description = db.Column(db.Text, default='')
    attack_type = db.Column(db.String(20), default='melee')  # melee, ranged

    matches = db.relationship('Match', backref='hero', lazy='dynamic')

    def get_winrate_for_user(self, user_id):
        total = self.matches.filter_by(user_id=user_id).count()
        if total == 0:
            return 0
        wins = self.matches.filter_by(user_id=user_id, result='win').count()
        return round(wins / total * 100, 1)

    def get_global_winrate(self):
        total = self.matches.count()
        if total == 0:
            return 0
        wins = self.matches.filter_by(result='win').count()
        return round(wins / total * 100, 1)

    def __repr__(self):
        return f'<Hero {self.name}>'


class Match(db.Model):
    __tablename__ = 'matches'

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False, index=True)
    hero_id = db.Column(db.Integer, db.ForeignKey('heroes.id'), nullable=False, index=True)
    match_type_id = db.Column(db.Integer, db.ForeignKey('match_types.id'), nullable=False)
    date_played = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    duration_minutes = db.Column(db.Integer, nullable=False)
    result = db.Column(db.String(10), nullable=False)  # win, loss
    kills = db.Column(db.Integer, default=0)
    deaths = db.Column(db.Integer, default=0)
    assists = db.Column(db.Integer, default=0)
    team_role = db.Column(db.String(30), default='')  # carry, mid, offlane, support, hard support
    gpm = db.Column(db.Integer, default=0)
    xpm = db.Column(db.Integer, default=0)
    notes = db.Column(db.Text, default='')

    @property
    def kda_ratio(self):
        return round((self.kills + self.assists) / max(self.deaths, 1), 2)

    @property
    def kda_string(self):
        return f'{self.kills}/{self.deaths}/{self.assists}'

    def __repr__(self):
        return f'<Match {self.id} by User {self.user_id}>'