from flask_wtf import FlaskForm
from wtforms import (StringField, PasswordField, SelectField, IntegerField,
                     TextAreaField, DateTimeLocalField, BooleanField, SubmitField)
from wtforms.validators import DataRequired, Email, Length, EqualTo, NumberRange, Optional, ValidationError
from models import User


class LoginForm(FlaskForm):
    username = StringField('Имя пользователя', validators=[DataRequired()])
    password = PasswordField('Пароль', validators=[DataRequired()])
    remember = BooleanField('Запомнить меня')
    submit = SubmitField('Войти')


class RegisterForm(FlaskForm):
    username = StringField('Имя пользователя',
                           validators=[DataRequired(), Length(min=3, max=64)])
    email = StringField('Email', validators=[DataRequired(), Email()])
    password = PasswordField('Пароль', validators=[DataRequired(), Length(min=6)])
    password2 = PasswordField('Повторите пароль',
                              validators=[DataRequired(), EqualTo('password',
                                                                   message='Пароли должны совпадать')])
    submit = SubmitField('Зарегистрироваться')

    def validate_username(self, field):
        if User.query.filter_by(username=field.data).first():
            raise ValidationError('Это имя пользователя уже занято.')

    def validate_email(self, field):
        if User.query.filter_by(email=field.data).first():
            raise ValidationError('Этот email уже зарегистрирован.')


class MatchForm(FlaskForm):
    hero_id = SelectField('Герой', coerce=int, validators=[DataRequired()])
    match_type_id = SelectField('Тип матча', coerce=int, validators=[DataRequired()])
    date_played = DateTimeLocalField('Дата и время', format='%Y-%m-%dT%H:%M',
                                     validators=[DataRequired()])
    duration_minutes = IntegerField('Длительность (мин)',
                                    validators=[DataRequired(), NumberRange(min=1, max=180)])
    result = SelectField('Результат', choices=[('win', 'Победа'), ('loss', 'Поражение')],
                         validators=[DataRequired()])
    kills = IntegerField('Убийства', validators=[DataRequired(), NumberRange(min=0)])
    deaths = IntegerField('Смерти', validators=[DataRequired(), NumberRange(min=0)])
    assists = IntegerField('Помощь', validators=[DataRequired(), NumberRange(min=0)])
    team_role = SelectField('Роль в команде', choices=[
        ('carry', 'Carry'), ('mid', 'Mid'), ('offlane', 'Offlane'),
        ('support', 'Support (4)'), ('hard_support', 'Hard Support (5)')
    ], validators=[DataRequired()])
    gpm = IntegerField('GPM', validators=[Optional(), NumberRange(min=0)])
    xpm = IntegerField('XPM', validators=[Optional(), NumberRange(min=0)])
    notes = TextAreaField('Заметки', validators=[Optional(), Length(max=500)])
    submit = SubmitField('Сохранить')


class HeroForm(FlaskForm):
    name = StringField('Название героя', validators=[DataRequired(), Length(max=100)])
    attribute = SelectField('Атрибут', choices=[
        ('strength', 'Сила'), ('agility', 'Ловкость'),
        ('intelligence', 'Интеллект'), ('universal', 'Универсальный')
    ], validators=[DataRequired()])
    role_id = SelectField('Роль', coerce=int, validators=[DataRequired()])
    difficulty = SelectField('Сложность', choices=[
        (1, '1 - Лёгкий'), (2, '2 - Средний'), (3, '3 - Сложный')
    ], coerce=int, validators=[DataRequired()])
    attack_type = SelectField('Тип атаки', choices=[
        ('melee', 'Ближний бой'), ('ranged', 'Дальний бой')
    ], validators=[DataRequired()])
    description = TextAreaField('Описание', validators=[Optional(), Length(max=1000)])
    submit = SubmitField('Сохранить')


class HeroRoleForm(FlaskForm):
    name = StringField('Название роли', validators=[DataRequired(), Length(max=50)])
    submit = SubmitField('Сохранить')


class MatchTypeForm(FlaskForm):
    name = StringField('Тип матча', validators=[DataRequired(), Length(max=50)])
    submit = SubmitField('Сохранить')


class UserEditForm(FlaskForm):
    username = StringField('Имя пользователя', validators=[DataRequired(), Length(max=64)])
    email = StringField('Email', validators=[DataRequired(), Email()])
    is_admin = BooleanField('Администратор')
    is_active_user = BooleanField('Активный')
    submit = SubmitField('Сохранить')