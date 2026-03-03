import logging
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardRemove, BotCommand, MenuButtonCommands
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler
from datetime import datetime
import psycopg2
from psycopg2.extras import RealDictCursor
import os
from io import BytesIO
import openpyxl
from openpyxl import Workbook

# Logging sozlash
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

# Bosqichlar
ISM, FAMILIYA, TELEFON, KANALLAR, MAQSAD, BROADCAST_MESSAGE, TAHRIRLASH_TANLOV, TAHRIRLASH_YANGI = range(8)

# ADMIN ID (o'zingizning Telegram ID ingizni kiriting!)
ADMIN_IDS = [5304449378, 6917103687]

# Tekshirish kerak bo'lgan kanallar
REQUIRED_CHANNELS = [
    "@ziyoaiuz",
    "@StarupAmbassadors"
]

# Kanal nomlarini ko'rsatish uchun
CHANNEL_NAMES = {
    "@ziyoaiuz": "ZiyoAI Community",
    "@StarupAmbassadors": "Startup Ambassadors | Kashkadarya"
}

# Database connection string
DATABASE_URL = os.getenv('DATABASE_URL')

# Startup da tekshirish
def check_environment():
    """Environment variablelarni tekshirish"""
    if not DATABASE_URL:
        logging.error("❌ DATABASE_URL topilmadi!")
        logging.error("Railway Variables da DATABASE_URL o'rnatilganligiga ishonch hosil qiling.")
        logging.error("PostgreSQL service reference qo'shilganmi tekshiring.")
        return False
    
    logging.info(f"✅ DATABASE_URL topildi: {DATABASE_URL[:20]}...")
    return True

# Database connection olish
def get_db_connection():
    """PostgreSQL ga ulanish"""
    try:
        if not DATABASE_URL:
            raise Exception("DATABASE_URL environment variable mavjud emas!")
        
        logging.info("📡 Database ga ulanish...")
        conn = psycopg2.connect(DATABASE_URL, connect_timeout=10)
        logging.info("✅ Database ga muvaffaqiyatli ulandi!")
        return conn
    except psycopg2.OperationalError as e:
        logging.error(f"❌ Database ga ulanishda xato (OperationalError): {e}")
        logging.error("DATABASE_URL to'g'ri formatda ekanligini tekshiring:")
        logging.error("postgresql://user:password@host:port/database")
        raise
    except Exception as e:
        logging.error(f"❌ Database ga ulanishda xato: {e}")
        raise

# Database yaratish va ulanish
def init_database():
    """Database jadvalini yaratish"""
    try:
        logging.info("🔧 Database ni sozlash...")
        conn = get_db_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                user_id BIGINT UNIQUE NOT NULL,
                username TEXT,
                ism TEXT NOT NULL,
                familiya TEXT NOT NULL,
                telefon TEXT NOT NULL,
                maqsad TEXT NOT NULL,
                sana TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        conn.commit()
        cursor.close()
        conn.close()
        logging.info("✅ Database tayyor!")
        return True
    except Exception as e:
        logging.error(f"❌ Database yaratishda xato: {e}")
        return False

# Database ga ma'lumot saqlash
def save_to_database(data):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            INSERT INTO users (user_id, username, ism, familiya, telefon, maqsad)
            VALUES (%s, %s, %s, %s, %s, %s)
        ''', (
            data['user_id'],
            data['username'],
            data['ism'],
            data['familiya'],
            data['telefon'],
            data['maqsad']
        ))
        
        conn.commit()
        
        # Tekshirish: haqiqatan ham qo'shildimi?
        cursor.execute('SELECT * FROM users WHERE user_id = %s', (data['user_id'],))
        result = cursor.fetchone()
        
        cursor.close()
        conn.close()
        
        if result:
            logging.info(f"✅ User {data['user_id']} muvaffaqiyatli saqlandi!")
            return True
        else:
            logging.error(f"❌ User {data['user_id']} saqlanmadi!")
            return False
            
    except psycopg2.IntegrityError as e:
        logging.error(f"User allaqachon mavjud: {e}")
        if conn:
            conn.rollback()
            conn.close()
        return False
    except Exception as e:
        logging.error(f"Database ga saqlashda xato: {e}")
        if conn:
            conn.rollback()
            conn.close()
        return False

# Database ma'lumotni yangilash
def update_user_data(user_id, field, value):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # SQL injection oldini olish uchun allowed fields
        allowed_fields = ['ism', 'familiya', 'telefon', 'maqsad']
        if field not in allowed_fields:
            return False
        
        query = f"UPDATE users SET {field} = %s WHERE user_id = %s"
        cursor.execute(query, (value, user_id))
        
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        logging.error(f"Ma'lumotni yangilashda xato: {e}")
        if conn:
            conn.rollback()
            conn.close()
        return False

# User ma'lumotlarini olish
def get_user_data(user_id):
    try:
        conn = get_db_connection()
        cursor = conn.cursor(cursor_factory=RealDictCursor)
        
        cursor.execute('SELECT * FROM users WHERE user_id = %s', (user_id,))
        row = cursor.fetchone()
        
        cursor.close()
        conn.close()
        
        if row:
            return dict(row)
        return None
    except Exception as e:
        logging.error(f"User ma'lumotlarini olishda xato: {e}")
        return None

# User ro'yxatdan o'tganmi tekshirish
def is_user_registered(user_id):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT id FROM users WHERE user_id = %s', (user_id,))
        result = cursor.fetchone()
        
        cursor.close()
        conn.close()
        return result is not None
    except Exception as e:
        logging.error(f"User tekshirishda xato: {e}")
        return False

# Barcha userlarni olish
def get_all_users():
    try:
        conn = get_db_connection()
        cursor = conn.cursor(cursor_factory=RealDictCursor)
        
        cursor.execute('SELECT * FROM users ORDER BY sana DESC')
        rows = cursor.fetchall()
        
        cursor.close()
        conn.close()
        
        return [dict(row) for row in rows]
    except Exception as e:
        logging.error(f"Barcha userlarni olishda xato: {e}")
        return []

# Statistika olish
def get_statistics():
    users = get_all_users()
    if not users:
        return "Hali hech kim ro'yxatdan o'tmagan"
    
    total = len(users)
    today = datetime.now().strftime('%Y-%m-%d')
    today_count = sum(1 for u in users if str(u['sana']).startswith(today))
    
    return f"""
📊 STATISTIKA

👥 Jami ro'yxatdan o'tganlar: {total} ta
📅 Bugun ro'yxatdan o'tganlar: {today_count} ta
⏰ Oxirgi ro'yxat: {users[0]['sana'] if users else 'Mavjud emas'}
    """

# Excel ga eksport qilish
def export_to_excel():
    users = get_all_users()
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Foydalanuvchilar"
    
    # Sarlavhalar
    ws.append(['№', 'Ism', 'Familiya', 'Telefon', 'Maqsad', 'Telegram ID', 'Username', 'Sana'])
    
    # Ma'lumotlar
    for user in users:
        ws.append([
            user['id'],
            user['ism'],
            user['familiya'],
            user['telefon'],
            user['maqsad'],
            user['user_id'],
            user['username'],
            str(user['sana'])
        ])
    
    # BytesIO ga saqlash (file system ga emas)
    excel_file = BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)
    
    return excel_file

# Admin tekshirish
def is_admin(user_id):
    return user_id in ADMIN_IDS

# Asosiy menyu
def get_main_menu():
    keyboard = [
        ["ℹ️ Ma'lumot", "📞 Bog'lanish"],
        ["✏️ Tahrirlash"]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

# Admin menyu
def get_admin_menu():
    keyboard = [
        ["📊 Statistika", "👥 Userlar"],
        ["📢 Xabar yuborish", "📁 Excel yuklab olish"],
        ["🔙 Oddiy menyu"]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

# /start buyrug'i
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    # Agar user allaqachon ro'yxatdan o'tgan bo'lsa
    if is_user_registered(user_id):
        if is_admin(user_id):
            menu = get_admin_menu()
            xabar = f"👋 Xush kelibsiz, Admin!\n\n📋 Admin panel ochiq."
        else:
            menu = get_main_menu()
            xabar = f"👋 Xush kelibsiz!\n\n✅ Siz allaqachon ro'yxatdan o'tgansiz!"
        
        await update.message.reply_text(xabar, reply_markup=menu)
        return ConversationHandler.END
    
    context.user_data.clear()
    
    await update.message.reply_text(
        f"Assalomu alaykum! 👋\n\n"
        "Ro'yxatdan o'tish uchun ismingizni kiriting:",
        reply_markup=ReplyKeyboardRemove()
    )
    return ISM

# Ism qabul qilish
async def ism_qabul(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['ism'] = update.message.text
    await update.message.reply_text(
        f"Yaxshi, {update.message.text}! Endi familiyangizni kiriting:"
    )
    return FAMILIYA

# Familiya qabul qilish
async def familiya_qabul(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['familiya'] = update.message.text
    
    keyboard = [[KeyboardButton("📱 Telefon raqamni yuborish", request_contact=True)]]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=True)
    
    await update.message.reply_text(
        "Ajoyib! Endi telefon raqamingizni yuboring:",
        reply_markup=reply_markup
    )
    return TELEFON

# Telefon raqam qabul qilish
async def telefon_qabul(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.contact:
        telefon = update.message.contact.phone_number
    else:
        telefon = update.message.text
    
    context.user_data['telefon'] = telefon
    user_id = update.effective_user.id
    
    # Kanallarga a'zolikni tekshirish
    azo_emas = []
    
    # TEST MODE: Adminlar uchun kanal tekshiruvini o'tkazib yuborish
    if is_admin(user_id):
        logging.info(f"Admin {user_id} - kanal tekshiruvi o'tkazib yuborildi")
    else:
        for kanal in REQUIRED_CHANNELS:
            try:
                member = await context.bot.get_chat_member(kanal, user_id)
                if member.status not in ['member', 'administrator', 'creator']:
                    azo_emas.append(kanal)
            except Exception as e:
                logging.error(f"Kanal tekshirishda xato {kanal}: {e}")
                azo_emas.append(kanal)
    
    if not azo_emas:
        await update.message.reply_text(
            "Ajoyib! Siz barcha kanallarga a'zosiz ✅\n\n"
            "Endi qatnashish maqsadingizni yozing:",
            reply_markup=ReplyKeyboardRemove()
        )
        return MAQSAD
    else:
        keyboard = []
        for kanal in azo_emas:
            kanal_nomi = CHANNEL_NAMES.get(kanal, kanal)
            if kanal.startswith('@'):
                url = f"https://t.me/{kanal[1:]}"
            else:
                url = f"https://t.me/{REQUIRED_CHANNELS[0][1:]}"
            keyboard.append([InlineKeyboardButton(f"A'zo bo'lish: {kanal_nomi}", url=url)])
        
        keyboard.append([InlineKeyboardButton("✅ Tekshirish", callback_data="check_membership")])
        reply_markup = InlineKeyboardMarkup(keyboard)
        
        await update.message.reply_text(
            "❌ Iltimos, avval quyidagi kanallarga a'zo bo'ling:\n\n" + 
            "\n".join([CHANNEL_NAMES.get(k, k) for k in azo_emas]) + 
            "\n\nA'zo bo'lganingizdan keyin 'Tekshirish' tugmasini bosing.",
            reply_markup=reply_markup
        )
        return KANALLAR

# Kanal a'zoligini qayta tekshirish
async def kanallarni_tekshir(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    user_id = update.effective_user.id
    azo_emas = []
    
    for kanal in REQUIRED_CHANNELS:
        try:
            member = await context.bot.get_chat_member(kanal, user_id)
            if member.status not in ['member', 'administrator', 'creator']:
                azo_emas.append(kanal)
        except Exception as e:
            logging.error(f"Kanal tekshirishda xato {kanal}: {e}")
            azo_emas.append(kanal)
    
    if not azo_emas:
        await query.edit_message_text(
            "✅ Ajoyib! Siz barcha kanallarga a'zosiz!\n\n"
            "Endi qatnashish maqsadingizni yozing:"
        )
        return MAQSAD
    else:
        keyboard = []
        for kanal in azo_emas:
            kanal_nomi = CHANNEL_NAMES.get(kanal, kanal)
            if kanal.startswith('@'):
                url = f"https://t.me/{kanal[1:]}"
            else:
                url = f"https://t.me/{REQUIRED_CHANNELS[0][1:]}"
            keyboard.append([InlineKeyboardButton(f"A'zo bo'lish: {kanal_nomi}", url=url)])
        
        keyboard.append([InlineKeyboardButton("✅ Tekshirish", callback_data="check_membership")])
        reply_markup = InlineKeyboardMarkup(keyboard)
        
        try:
            await query.edit_message_text(
                "❌ Hali ham quyidagi kanallarga a'zo emassiz:\n\n" + 
                "\n".join([CHANNEL_NAMES.get(k, k) for k in azo_emas]) + 
                "\n\nIltimos, a'zo bo'ling va qayta tekshiring.",
                reply_markup=reply_markup
            )
        except Exception:
            pass
        return KANALLAR

# Maqsadni qabul qilish va yakunlash
async def maqsad_qabul(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['maqsad'] = update.message.text
    context.user_data['user_id'] = update.effective_user.id
    context.user_data['username'] = update.effective_user.username or 'Username yo\'q'
    
    # Database ga saqlash
    if save_to_database(context.user_data):
        save_status = "\n💾 Ma'lumotlar saqlandi!"
    else:
        save_status = "\n⚠️ Ma'lumotlarni saqlashda xatolik yuz berdi"
    
    # User ga javob
    malumot = f"""
✅ Ro'yxatdan o'tish muvaffaqiyatli yakunlandi!

📝 Sizning ma'lumotlaringiz:
👤 Ism: {context.user_data.get('ism')}
👤 Familiya: {context.user_data.get('familiya')}
📱 Telefon: {context.user_data.get('telefon')}
🎯 Maqsad: {context.user_data.get('maqsad')}
{save_status}

📢 **Yangi Eventlarni, Workshoplarni va Yangililarni o'tqaib yubormasligiz uchun tepadagi kanallarni kuzatishda davom eting**:
https://t.me/StarupAmbassadors

📢 Klubimizga qo'shiling: https://startup-ambassadors-registration-production.up.railway.app/

Rahmat! ✨
    """
    
    await update.message.reply_text(malumot, reply_markup=get_main_menu())
    
    # Adminga xabar yuborish
    admin_notification = f"""
🆕 YANGI RO'YXAT!

👤 Ism: {context.user_data.get('ism')} {context.user_data.get('familiya')}
📱 Telefon: {context.user_data.get('telefon')}
🎯 Maqsad: {context.user_data.get('maqsad')}
🆔 Telegram ID: {context.user_data.get('user_id')}
👤 Username: @{context.user_data.get('username')}
⏰ Vaqt: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
    """
    
    for admin_id in ADMIN_IDS:
        try:
            await context.bot.send_message(chat_id=admin_id, text=admin_notification)
        except Exception as e:
            logging.error(f"Adminga xabar yuborishda xato: {e}")
    
    context.user_data.clear()
    return ConversationHandler.END

# TAHRIRLASH funksiyasi
async def tahrirlash_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if not is_user_registered(user_id):
        await update.message.reply_text(
            "❌ Siz hali ro'yxatdan o'tmagansiz!\n"
            "Iltimos, /start buyrug'ini bosing.",
            reply_markup=ReplyKeyboardRemove()
        )
        return ConversationHandler.END
    
    user_data = get_user_data(user_id)
    
    keyboard = [
        ["✏️ Ismni tahrirlash", "✏️ Familiyani tahrirlash"],
        ["✏️ Telefonni tahrirlash", "✏️ Maqsadni tahrirlash"],
        ["❌ Bekor qilish"]
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    
    await update.message.reply_text(
        f"✏️ TAHRIRLASH\n\n"
        f"Hozirgi ma'lumotlaringiz:\n"
        f"👤 Ism: {user_data['ism']}\n"
        f"👤 Familiya: {user_data['familiya']}\n"
        f"📱 Telefon: {user_data['telefon']}\n"
        f"🎯 Maqsad: {user_data['maqsad']}\n\n"
        f"Nimani tahrirlashni xohlaysiz?",
        reply_markup=reply_markup
    )
    return TAHRIRLASH_TANLOV

async def tahrirlash_tanlov(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    
    if text == "❌ Bekor qilish":
        return await bekor(update, context)
    
    field_map = {
        "✏️ Ismni tahrirlash": ("ism", "Yangi ismingizni kiriting:"),
        "✏️ Familiyani tahrirlash": ("familiya", "Yangi familiyangizni kiriting:"),
        "✏️ Telefonni tahrirlash": ("telefon", "Yangi telefon raqamingizni kiriting:"),
        "✏️ Maqsadni tahrirlash": ("maqsad", "Yangi maqsadingizni kiriting:")
    }
    
    if text in field_map:
        context.user_data['tahrirlash_field'] = field_map[text][0]
        await update.message.reply_text(
            field_map[text][1],
            reply_markup=ReplyKeyboardRemove()
        )
        return TAHRIRLASH_YANGI
    
    await update.message.reply_text("❌ Noto'g'ri tanlov. Qaytadan urinib ko'ring.")
    return TAHRIRLASH_TANLOV

async def tahrirlash_yangi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    field = context.user_data.get('tahrirlash_field')
    new_value = update.message.text
    
    if update_user_data(user_id, field, new_value):
        field_names = {
            'ism': 'Ism',
            'familiya': 'Familiya',
            'telefon': 'Telefon',
            'maqsad': 'Maqsad'
        }
        
        await update.message.reply_text(
            f"✅ {field_names.get(field)} muvaffaqiyatli yangilandi!\n\n"
            f"Yangi qiymat: {new_value}",
            reply_markup=get_main_menu()
        )
    else:
        await update.message.reply_text(
            "❌ Ma'lumotni yangilashda xatolik yuz berdi.",
            reply_markup=get_main_menu()
        )
    
    context.user_data.clear()
    return ConversationHandler.END

# MENU TUGMALARI HANDLERI
async def handle_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    user_id = update.effective_user.id
    
    # Admin tugmalari
    if is_admin(user_id):
        if text == "📊 Statistika":
            await stats_command(update, context)
            return
        elif text == "👥 Userlar":
            await users_command(update, context)
            return
        elif text == "📢 Xabar yuborish":
            await broadcast_command(update, context)
            return BROADCAST_MESSAGE
        elif text == "📁 Excel yuklab olish":
            try:
                excel_file = export_to_excel()
                await update.message.reply_document(
                    document=excel_file,
                    filename=f'users_data_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
                    caption="📊 Barcha foydalanuvchilar ma'lumotlari"
                )
            except Exception as e:
                logging.error(f"Excel eksport xatosi: {e}")
                await update.message.reply_text("Excel faylni yaratishda xatolik yuz berdi")
            return
        elif text == "🔙 Oddiy menyu":
            await update.message.reply_text(
                "Oddiy menyu:",
                reply_markup=get_main_menu()
            )
            return
    
    # Oddiy user tugmalari
    if text == "ℹ️ Ma'lumot":
        await update.message.reply_text(
            "ℹ️ *BIZ HAQIMIZDA*\n\n"
            "Ushbu bot *start-up loyihalarini rivojlantirish va startuperlarni bog'lash* maqsadida yaratildi\n"
            "Bizning kanallarimizga a'zo bo'ling, *oxirgi yangiliklar, foydali resurslar va startuperlar uchun imkoniyatlar* bilan tanishing! 🚀\n\n"
            "📢 Asosiy kanal: https://t.me/StarupAmbassadors",
            parse_mode='Markdown'
        )
    elif text == "📞 Bog'lanish":
        await update.message.reply_text(
            "📞 BOG'LANISH\n\n"
            "Savollaringiz bo'lsa, biz bilan bog'laning:\n"
            "📧 Email: nematovm764@gmail.com\n"
            "📱 Telefon: +998 91 642 81 86"
        )
    elif text == "✏️ Tahrirlash":
        await tahrirlash_command(update, context)
        return TAHRIRLASH_TANLOV
    elif text == "📊 Statistika" and not is_admin(user_id):
        await update.message.reply_text(
            "❌ Bu bo'lim faqat adminlar uchun!"
        )

# ADMIN BUYRUQLARI
async def stats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        await update.message.reply_text("❌ Bu buyruq faqat adminlar uchun!")
        return
    
    stats = get_statistics()
    menu = get_admin_menu() if is_admin(user_id) else None
    await update.message.reply_text(stats, reply_markup=menu)

async def users_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        await update.message.reply_text("❌ Bu buyruq faqat adminlar uchun!")
        return
    
    users = get_all_users()
    
    if not users:
        await update.message.reply_text("Hali hech kim ro'yxatdan o'tmagan", reply_markup=get_admin_menu())
        return
    
    # Oxirgi 10 ta userni ko'rsatish
    last_users = users[:10]
    message = "👥 OXIRGI 10 TA USER:\n\n"
    
    for user in last_users:
        message += f"#{user['id']} {user['ism']} {user['familiya']}\n"
        message += f"📱 {user['telefon']}\n"
        message += f"🆔 {user['user_id']}\n"
        message += f"📅 {user['sana']}\n\n"
    
    await update.message.reply_text(message, reply_markup=get_admin_menu())

async def broadcast_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        await update.message.reply_text("❌ Bu buyruq faqat adminlar uchun!")
        return ConversationHandler.END
    
    await update.message.reply_text(
        "📢 Barcha foydalanuvchilarga yubormoqchi bo'lgan xabaringizni yozing:\n\n"
        "Bekor qilish uchun /cancel ni yozing.",
        reply_markup=ReplyKeyboardRemove()
    )
    return BROADCAST_MESSAGE

async def send_broadcast(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message_text = update.message.text
    users = get_all_users()
    
    if not users:
        await update.message.reply_text("Hali hech kim ro'yxatdan o'tmagan")
        return ConversationHandler.END
    
    success_count = 0
    fail_count = 0
    
    status_msg = await update.message.reply_text(f"📤 Yuborilmoqda... 0/{len(users)}")
    
    for i, user in enumerate(users):
        try:
            await context.bot.send_message(chat_id=user['user_id'], text=message_text)
            success_count += 1
        except Exception as e:
            logging.error(f"User {user['user_id']} ga yuborishda xato: {e}")
            fail_count += 1
        
        # Har 5 ta userdan keyin statusni yangilash
        if (i + 1) % 5 == 0:
            await status_msg.edit_text(f"📤 Yuborilmoqda... {i+1}/{len(users)}")
    
    await status_msg.edit_text(
        f"✅ Xabar yuborish tugadi!\n\n"
        f"✅ Muvaffaqiyatli: {success_count}\n"
        f"❌ Xato: {fail_count}\n"
        f"📊 Jami: {len(users)}"
    )
    
    await update.message.reply_text(
        "Admin panel:",
        reply_markup=get_admin_menu()
    )
    
    return ConversationHandler.END

async def bekor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if is_user_registered(user_id):
        menu = get_admin_menu() if is_admin(user_id) else get_main_menu()
        await update.message.reply_text("Jarayon bekor qilindi.", reply_markup=menu)
    else:
        await update.message.reply_text(
            "Jarayon bekor qilindi. /start ni bosing qaytadan boshlash uchun.",
            reply_markup=ReplyKeyboardRemove()
        )
    
    context.user_data.clear()
    return ConversationHandler.END

def main():
    # Environment tekshirish
    if not check_environment():
        logging.error("❌ Environment variableler to'g'ri sozlanmagan! Bot ishga tushmaydi.")
        return
    
    # TOKEN ni environment variable dan olish
    TOKEN = os.getenv('BOT_TOKEN')
    
    if not TOKEN:
        logging.error("❌ BOT_TOKEN topilmadi! Railway Variables da BOT_TOKEN o'rnatilganligiga ishonch hosil qiling.")
        return
    
    logging.info(f"✅ BOT_TOKEN topildi: {TOKEN[:10]}...")
    
    # Database ni boshlash
    if not init_database():
        logging.error("❌ Database ishga tushmadi! Iltimos, PostgreSQL service ishga tushganligini tekshiring.")
        return
    
    application = Application.builder().token(TOKEN).build()
    
    # Bot buyruqlarini va Menu tugmasini sozlash
    async def post_init(app: Application):
        commands = [
            BotCommand("start", "🏠 Botni ishga tushirish"),
            BotCommand("stats", "📊 Statistika (Admin)"),
            BotCommand("users", "👥 Foydalanuvchilar ro'yxati (Admin)"),
            BotCommand("broadcast", "📢 Xabar yuborish (Admin)"),
            BotCommand("cancel", "❌ Jarayonni bekor qilish"),
        ]
        await app.bot.set_my_commands(commands)
        await app.bot.set_chat_menu_button(menu_button=MenuButtonCommands())
    
    application.post_init = post_init
    
    # Asosiy suhbat
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('start', start)],
        states={
            ISM: [MessageHandler(filters.TEXT & ~filters.COMMAND, ism_qabul)],
            FAMILIYA: [MessageHandler(filters.TEXT & ~filters.COMMAND, familiya_qabul)],
            TELEFON: [MessageHandler((filters.CONTACT | filters.TEXT) & ~filters.COMMAND, telefon_qabul)],
            KANALLAR: [CallbackQueryHandler(kanallarni_tekshir, pattern="check_membership")],
            MAQSAD: [MessageHandler(filters.TEXT & ~filters.COMMAND, maqsad_qabul)],
        },
        fallbacks=[CommandHandler('cancel', bekor)],
        allow_reentry=True
    )
    
    # Tahrirlash suhbati
    tahrirlash_handler = ConversationHandler(
        entry_points=[
            MessageHandler(filters.Regex('^✏️ Tahrirlash$'), tahrirlash_command)
        ],
        states={
            TAHRIRLASH_TANLOV: [MessageHandler(filters.TEXT & ~filters.COMMAND, tahrirlash_tanlov)],
            TAHRIRLASH_YANGI: [MessageHandler(filters.TEXT & ~filters.COMMAND, tahrirlash_yangi)],
        },
        fallbacks=[CommandHandler('cancel', bekor)]
    )
    
    # Broadcast suhbati
    broadcast_handler = ConversationHandler(
        entry_points=[
            CommandHandler('broadcast', broadcast_command),
            MessageHandler(filters.Regex('^📢 Xabar yuborish$'), broadcast_command)
        ],
        states={
            BROADCAST_MESSAGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, send_broadcast)],
        },
        fallbacks=[CommandHandler('cancel', bekor)]
    )
    
    application.add_handler(conv_handler)
    application.add_handler(tahrirlash_handler)
    application.add_handler(broadcast_handler)
    application.add_handler(CommandHandler('stats', stats_command))
    application.add_handler(CommandHandler('users', users_command))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_menu))
    
    logging.info("🚀 Bot ishga tushdi!")
    logging.info("💾 Database: PostgreSQL")
    logging.info("📋 Menu tugmasi qo'shildi!")
    
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == '__main__':
    main()