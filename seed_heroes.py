"""
Запуск: python seed_heroes.py
Заполняет БД базовым набором героев Dota 2
"""
from app import app, db
from models import Hero, HeroRole

heroes_data = [
    # (name, attribute, role_name, difficulty, attack_type)
    ('Anti-Mage', 'agility', 'Carry', 1, 'melee'),
    ('Axe', 'strength', 'Initiator', 1, 'melee'),
    ('Bane', 'intelligence', 'Disabler', 2, 'ranged'),
    ('Bloodseeker', 'agility', 'Carry', 1, 'melee'),
    ('Crystal Maiden', 'intelligence', 'Support', 1, 'ranged'),
    ('Drow Ranger', 'agility', 'Carry', 1, 'ranged'),
    ('Earthshaker', 'strength', 'Initiator', 2, 'melee'),
    ('Juggernaut', 'agility', 'Carry', 1, 'melee'),
    ('Mirana', 'universal', 'Nuker', 2, 'ranged'),
    ('Morphling', 'agility', 'Carry', 3, 'ranged'),
    ('Shadow Fiend', 'agility', 'Nuker', 2, 'ranged'),
    ('Phantom Lancer', 'agility', 'Carry', 2, 'melee'),
    ('Puck', 'intelligence', 'Initiator', 3, 'ranged'),
    ('Pudge', 'strength', 'Durable', 2, 'melee'),
    ('Razor', 'agility', 'Carry', 1, 'ranged'),
    ('Sand King', 'strength', 'Initiator', 2, 'melee'),
    ('Storm Spirit', 'intelligence', 'Nuker', 3, 'ranged'),
    ('Sven', 'strength', 'Carry', 1, 'melee'),
    ('Tiny', 'strength', 'Durable', 2, 'melee'),
    ('Vengeful Spirit', 'agility', 'Support', 1, 'ranged'),
    ('Windranger', 'intelligence', 'Nuker', 2, 'ranged'),
    ('Zeus', 'intelligence', 'Nuker', 1, 'ranged'),
    ('Kunkka', 'strength', 'Carry', 2, 'melee'),
    ('Lina', 'intelligence', 'Nuker', 2, 'ranged'),
    ('Lion', 'intelligence', 'Disabler', 1, 'ranged'),
    ('Shadow Shaman', 'intelligence', 'Pusher', 1, 'ranged'),
    ('Witch Doctor', 'intelligence', 'Support', 1, 'ranged'),
    ('Lich', 'intelligence', 'Support', 1, 'ranged'),
    ('Faceless Void', 'agility', 'Carry', 3, 'melee'),
    ('Wraith King', 'strength', 'Carry', 1, 'melee'),
    ('Phantom Assassin', 'agility', 'Carry', 1, 'melee'),
    ('Templar Assassin', 'agility', 'Carry', 2, 'ranged'),
    ('Viper', 'agility', 'Carry', 1, 'ranged'),
    ('Luna', 'agility', 'Carry', 1, 'ranged'),
    ('Dragon Knight', 'strength', 'Durable', 1, 'melee'),
    ('Dazzle', 'intelligence', 'Support', 1, 'ranged'),
    ('Clockwerk', 'universal', 'Initiator', 2, 'melee'),
    ('Huskar', 'strength', 'Carry', 2, 'ranged'),
    ('Night Stalker', 'strength', 'Initiator', 1, 'melee'),
    ('Broodmother', 'universal', 'Pusher', 2, 'melee'),
    ('Bounty Hunter', 'agility', 'Escape', 1, 'melee'),
    ('Weaver', 'agility', 'Carry', 2, 'ranged'),
    ('Jakiro', 'intelligence', 'Support', 1, 'ranged'),
    ('Batrider', 'universal', 'Initiator', 2, 'ranged'),
    ('Chen', 'intelligence', 'Support', 3, 'ranged'),
    ('Spectre', 'agility', 'Carry', 2, 'melee'),
    ('Doom', 'strength', 'Durable', 2, 'melee'),
    ('Ancient Apparition', 'intelligence', 'Support', 2, 'ranged'),
    ('Ursa', 'agility', 'Carry', 1, 'melee'),
    ('Spirit Breaker', 'strength', 'Initiator', 1, 'melee'),
    ('Invoker', 'intelligence', 'Nuker', 3, 'ranged'),
    ('Silencer', 'intelligence', 'Disabler', 2, 'ranged'),
    ('Outworld Destroyer', 'intelligence', 'Nuker', 2, 'ranged'),
    ('Treant Protector', 'strength', 'Support', 2, 'melee'),
    ('Ogre Magi', 'strength', 'Support', 1, 'melee'),
    ('Undying', 'strength', 'Durable', 1, 'melee'),
    ('Rubick', 'intelligence', 'Support', 3, 'ranged'),
    ('Disruptor', 'intelligence', 'Disabler', 2, 'ranged'),
    ('Nyx Assassin', 'agility', 'Initiator', 2, 'melee'),
    ('Naga Siren', 'agility', 'Carry', 3, 'melee'),
    ('Keeper of the Light', 'intelligence', 'Support', 2, 'ranged'),
    ('Io', 'strength', 'Support', 3, 'ranged'),
    ('Visage', 'intelligence', 'Nuker', 3, 'ranged'),
    ('Slark', 'agility', 'Carry', 2, 'melee'),
    ('Medusa', 'agility', 'Carry', 2, 'ranged'),
    ('Troll Warlord', 'agility', 'Carry', 2, 'melee'),
    ('Centaur Warrunner', 'strength', 'Durable', 1, 'melee'),
    ('Magnus', 'strength', 'Initiator', 2, 'melee'),
    ('Timbersaw', 'strength', 'Nuker', 2, 'melee'),
    ('Bristleback', 'strength', 'Durable', 1, 'melee'),
    ('Tusk', 'strength', 'Initiator', 2, 'melee'),
    ('Skywrath Mage', 'intelligence', 'Nuker', 1, 'ranged'),
    ('Abaddon', 'strength', 'Support', 1, 'melee'),
    ('Elder Titan', 'strength', 'Initiator', 2, 'melee'),
    ('Legion Commander', 'strength', 'Carry', 1, 'melee'),
    ('Ember Spirit', 'agility', 'Carry', 2, 'melee'),
    ('Earth Spirit', 'strength', 'Initiator', 3, 'melee'),
    ('Terrorblade', 'agility', 'Carry', 2, 'melee'),
    ('Phoenix', 'strength', 'Nuker', 2, 'ranged'),
    ('Oracle', 'intelligence', 'Support', 3, 'ranged'),
    ('Techies', 'intelligence', 'Nuker', 2, 'ranged'),
    ('Winter Wyvern', 'intelligence', 'Support', 2, 'ranged'),
    ('Arc Warden', 'agility', 'Carry', 3, 'ranged'),
    ('Underlord', 'strength', 'Durable', 1, 'melee'),
    ('Monkey King', 'agility', 'Carry', 2, 'melee'),
    ('Dark Willow', 'intelligence', 'Support', 2, 'ranged'),
    ('Pangolier', 'universal', 'Initiator', 2, 'melee'),
    ('Grimstroke', 'intelligence', 'Support', 2, 'ranged'),
    ('Mars', 'strength', 'Initiator', 1, 'melee'),
    ('Snapfire', 'strength', 'Support', 1, 'ranged'),
    ('Void Spirit', 'universal', 'Nuker', 2, 'melee'),
    ('Hoodwink', 'agility', 'Nuker', 2, 'ranged'),
    ('Dawnbreaker', 'strength', 'Durable', 1, 'melee'),
    ('Marci', 'universal', 'Carry', 1, 'melee'),
    ('Primal Beast', 'strength', 'Initiator', 1, 'melee'),
    ('Muerta', 'intelligence', 'Carry', 2, 'ranged'),
]

def seed():
    with app.app_context():
        if Hero.query.count() > 0:
            print(f'В БД уже есть {Hero.query.count()} героев. Пропускаем.')
            return

        roles_map = {}
        for role in HeroRole.query.all():
            roles_map[role.name] = role.id

        added = 0
        for name, attr, role_name, diff, atk in heroes_data:
            role_id = roles_map.get(role_name)
            if not role_id:
                print(f'Роль "{role_name}" не найдена, пропускаем {name}')
                continue
            hero = Hero(
                name=name,
                attribute=attr,
                role_id=role_id,
                difficulty=diff,
                attack_type=atk,
                description=f'{name} — герой Dota 2'
            )
            db.session.add(hero)
            added += 1

        db.session.commit()
        print(f'Добавлено {added} героев!')


if __name__ == '__main__':
    seed()