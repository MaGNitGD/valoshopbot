import sqlite3 as sq

db = sq.connect('valousers.db')
cur = db.cursor()

async def db_start():
    cur.execute('CREATE TABLE IF NOT EXISTS usersInfo('
                'id INTEGER PRIMARY KEY AUTOINCREMENT, '
                'tgid INTEGER,'
                'regdate DATE,'
                'balance INTEGER,'
                'fullinv TEXT,'
                'changingid TEXT,'
                'store TEXT)')
    cur.execute('CREATE TABLE IF NOT EXISTS usersEquipped('
                'id TEXT,'
                'hp INTEGER,'
                'dmg INTEGER,'
                'dodge INTEGER,'
                'armor INTEGER,'
                'headshot INTEGER,'
                'accuracy INTEGER,'
                'dodge_perc INTEGER,'
                'accuracy_perc INTEGER,'
                'dmg_plus INTEGER,'
                'armor_plus INTEGER,'
                'headshot_perc INTEGER,'
                'dodge_plus INTEGER,'
                'hp_plus INTEGER,'
                'hp_perc INTEGER,'
                'headshot_plus INTEGER,'
                'dmg_perc INTEGER,'
                'armor_perc INTEGER,'
                'accuracy_plus INTEGER)')
    db.commit()

async def cmdStart(user_id):
    user = cur.execute('SELECT * FROM usersInfo WHERE tgid == {key}'.format(key=user_id)).fetchone()
    if not user:
        cur.execute('INSERT INTO usersInfo (tgid) VALUES ({key})'.format(key=user_id))
        db.commit()