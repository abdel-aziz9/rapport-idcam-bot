"""
================================================================
  BOT TELEGRAM — RAPPORT JOURNALIER IDCAM
  Technicien : ABDOUL-AZIZ MOHAMADOU LAMINOU
  Version corrigée — Compatible Python 3.14 / PTB 21.x
================================================================
"""

import os
import datetime
import random
import warnings
import logging

warnings.filterwarnings("ignore")
logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, CommandHandler, CallbackQueryHandler, ContextTypes
)
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

# ─────────────────────────────────────────────
#  ⚙️  CONFIGURATION
# ─────────────────────────────────────────────

TOKEN_BOT     = os.environ.get("TOKEN_BOT", "8871644850:AAGqLgYPVo8ABS0xFtAhJJDoaa3_XdiN7Sk")
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template_rapport.xlsx")
DOSSIER       = os.path.dirname(__file__)

# ─────────────────────────────────────────────
#  DONNÉES INTERVENTIONS
# ─────────────────────────────────────────────

POSTES        = [f"poste {i:02d}" for i in range(1, 21)]
CENTRES_DIST  = ["NO27 DJAMBOUTOU","NO29 YELWA","NO06 LAGO","NO03 FIGUIL","NO08 GUIDER"]

POOL = [
    ("Logiciel","Enrôlement","local","camera scan doc - ecran noir",
     lambda p,c: f"Le {p} (centre {c}): camera scan doc affiche ecran noir. Reinitialisation camera et redemarrage machine. Resolu."),
    ("Logiciel","Enrôlement","local","camera figee sur ecran blanc",
     lambda p,c: f"Camera du {p} figee sur ecran blanc. Redemarrage complet effectue. Resolu."),
    ("Logiciel","Enrôlement","local","probleme reseau - camera et empreinte HS",
     lambda p,c: f"Le {p}: panne reseau rendant camera et empreinte inactives. Verification cable, test ping, redemarrage. Resolu."),
    ("Logiciel","Enrôlement","local","empreinte digitale non fonctionnelle",
     lambda p,c: f"Le {p}: empreinte dysfonctionnelle. Debranchement/rebranchement lecteur biometrique et redemarrage. Resolu."),
    ("Logiciel","Enrôlement","local","QR code non detecte par camera doc",
     lambda p,c: f"Le {p}: camera scan doc ne detecte pas le QR code. Retelechargement document et nouvelle tentative. Resolu."),
    ("Logiciel","Enrôlement","local","ajustement photo - personne agee ou handicapee",
     lambda p,c: f"Le {p}: difficulte prise de photo usager age. Assistance enrolleur, reglage luminosite et cadrage. Photo validee."),
    ("Logiciel","Enrôlement","local","logiciel QOMITTO ne s'ouvre pas",
     lambda p,c: f"Le {p}: QOMITTO ne s'ouvre pas. Redemarrage service et relance application. Resolu."),
    ("Logiciel","Enrôlement","local","redemarrage correctif - poste lent",
     lambda p,c: f"Le {p} presentait des lenteurs. Redemarrage complet effectue. Performances retablies."),
    ("Logiciel","Enrôlement","local","empreinte et camera inactives apres coupure reseau",
     lambda p,c: f"Coupure reseau sur {p}: empreinte et camera inactives. Redemarrage apres retablissement reseau. Resolu."),
    ("Logiciel","Enrôlement","distance","assistance distance - probleme reseau",
     lambda p,c: f"Assistance telephonique centre {c}: probleme reseau. Guidage test ping et redemarrage. Resolu a distance."),
    ("Logiciel","Enrôlement","distance","assistance distance - empreinte non reconnue",
     lambda p,c: f"Assistance a distance centre {c}: empreinte non reconnue. Instructions debranchement/rebranchement. Resolu."),
    ("Logiciel","Enrôlement","distance","assistance distance - camera bloquee",
     lambda p,c: f"Assistance telephonique centre {c}: camera bloquee. Guidage reinitialisation camera. Resolu a distance."),
    ("Logiciel","Enrôlement","distance","assistance distance - QOMITTO ne fonctionne pas",
     lambda p,c: f"Assistance a distance centre {c}: QOMITTO ne fonctionne pas. Guidage redemarrage service. Resolu."),
]

# ─────────────────────────────────────────────
#  GÉNÉRATION EXCEL
# ─────────────────────────────────────────────

def pick_interventions(n):
    ivs, seen = [], set()
    for _ in range(n * 10):
        if len(ivs) >= n: break
        typ, ss, loc, titre, desc_fn = random.choice(POOL)
        if titre in seen: continue
        seen.add(titre)
        if loc == "distance":
            c = random.choice(CENTRES_DIST)
            ivs.append({"type":typ,"ss":ss,"code":c,"titre":titre,"desc":desc_fn("",c)})
        else:
            p = random.choice(POSTES)
            ivs.append({"type":typ,"ss":ss,"code":"NO01BIS","titre":titre,"desc":desc_fn(p,"NO01BIS")})
    return ivs

def rand_times(n):
    slots, used = [], set()
    for _ in range(n * 20):
        if len(slots) >= n: break
        sm = random.randint(8*60+30, 15*60)
        k = sm // 15
        if k in used: continue
        used.add(k)
        em = min(sm + random.randint(10, 45), 16*60)
        slots.append((sm/1440, em/1440))
    return sorted(slots)

def to_time(frac):
    s = int(round(frac * 86400))
    h, r = divmod(s, 3600); m, s = divmod(r, 60)
    return datetime.time(h, m, s)

def generer_rapport(target_date):
    random.seed(int(target_date.strftime("%Y%m%d")))
    n   = random.randint(3, 6)
    ivs = pick_interventions(n)
    tms = rand_times(len(ivs))
    tot = sum(e - s for s, e in tms)

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb["Rapport Journalier"]

    ws["B2"] = target_date
    ws["B2"].number_format = "dd/mm/yyyy"

    for r in range(6, 9):
        for col in "CDEFG": ws[f"{col}{r}"] = None
    ws["D6"] = sum(1 for iv in ivs if iv["type"] == "Logiciel")

    ws["B13"] = "ABDOUL-AZIZ M.L"
    dv = DataValidation(type="list", formula1="Liste!$B$2:$B$34", showDropDown=False)
    dv.sqref = "B13"
    ws.add_data_validation(dv)
    ws["C13"] = len(ivs)
    ws["D13"] = to_time(tot)
    ws["D13"].number_format = "h:mm:ss"

    for r in range(18, 24):
        for col in "ABCDEFGHIJK": ws[f"{col}{r}"].value = None
    tpl = 4
    if len(ivs) > tpl: ws.insert_rows(22, len(ivs)-tpl)
    elif len(ivs) < tpl: ws.delete_rows(18+len(ivs), tpl-len(ivs))

    thin = Side(style="thin", color="FF000000")
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)
    dv2  = DataValidation(type="list", formula1="Liste!$B$2:$B$34", showDropDown=False)
    dv2.sqref = f"B18:B{17+len(ivs)}"
    ws.add_data_validation(dv2)

    for i, (iv, (ts, te)) in enumerate(zip(ivs, tms)):
        r = 18 + i
        def sc(col, val, fmt=None, center=False):
            cell = ws[f"{col}{r}"]
            cell.value = val
            cell.font  = Font(name="Calibri", size=11)
            cell.border = brd
            cell.alignment = Alignment(
                horizontal="center" if center else "left",
                vertical="center" if center else "top",
                wrap_text=True)
            if fmt: cell.number_format = fmt
        sc("A", i+1, center=True); sc("B","ABDOUL-AZIZ M.L")
        sc("C",iv["type"]);         sc("D",iv["ss"])
        sc("E",iv["code"],center=True); sc("F",iv["titre"]); sc("G",iv["desc"])
        sc("H","",center=True)
        sc("I",to_time(ts),fmt="h:mm",center=True)
        sc("J","Terminé",center=True)
        sc("K",to_time(te),fmt="h:mm",center=True)
        ws.row_dimensions[r].height = 57.0

    nom    = f"Rapport_ABDOUL-AZIZ_{target_date.strftime('%Y-%m-%d')}.xlsx"
    chemin = os.path.join(DOSSIER, nom)
    wb.save(chemin)
    return chemin, ivs, to_time(tot)

# ─────────────────────────────────────────────
#  ENVOI DU RAPPORT
# ─────────────────────────────────────────────

async def _envoyer(chat_or_query, date):
    # Récupérer l'objet message correct
    if hasattr(chat_or_query, "message"):
        reply = chat_or_query.message.reply_text
        send_doc = chat_or_query.message.reply_document
    else:
        reply = chat_or_query.reply_text
        send_doc = chat_or_query.reply_document

    msg = await reply("⏳ Génération du rapport en cours...")
    try:
        chemin, ivs, duree = generer_rapport(date)
        lignes = "\n".join(f"  {i+1}. [{iv['code']}] {iv['titre']}" for i,iv in enumerate(ivs))
        caption = (
            f"✅ *Rapport du {date.strftime('%d/%m/%Y')}*\n\n"
            f"👤 ABDOUL-AZIZ M.L | 🏢 NO01BIS\n"
            f"📊 *{len(ivs)} interventions* | ⏱ *{duree}*\n\n"
            f"*Détail :*\n{lignes}"
        )
        await msg.delete()
        with open(chemin, "rb") as f:
            await send_doc(document=f, filename=os.path.basename(chemin),
                           caption=caption, parse_mode="Markdown")
        os.remove(chemin)
    except FileNotFoundError:
        await msg.edit_text("❌ template_rapport.xlsx introuvable dans le dossier du bot.")
    except Exception as e:
        await msg.edit_text(f"❌ Erreur : {e}")

# ─────────────────────────────────────────────
#  COMMANDES
# ─────────────────────────────────────────────

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clavier = InlineKeyboardMarkup([
        [InlineKeyboardButton("📄 Rapport aujourd'hui", callback_data="today")],
        [InlineKeyboardButton("📅 Rapport hier",        callback_data="hier")],
        [InlineKeyboardButton("🕒 Activer envoi 15h",   callback_data="auto_on")],
        [InlineKeyboardButton("❓ Aide",                callback_data="aide")],
    ])
    await update.message.reply_text(
        "👋 *Bonjour ABDOUL-AZIZ !*\n\n"
        "Je génère votre rapport journalier IDCAM automatiquement.\n\n"
        "Choisissez une option 👇",
        parse_mode="Markdown", reply_markup=clavier
    )

async def cmd_rapport(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await _envoyer(update, datetime.date.today())

async def cmd_hier(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await _envoyer(update, datetime.date.today() - datetime.timedelta(days=1))

async def cmd_aide(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📖 *AIDE — Bot Rapport IDCAM*\n\n"
        "• /rapport → Rapport d'aujourd'hui\n"
        "• /hier    → Rapport d'hier\n"
        "• /auto    → Envoi automatique 15h\n"
        "• /stop    → Arrêter l'automatique\n"
        "• /start   → Menu principal",
        parse_mode="Markdown"
    )

async def cmd_auto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    for job in context.job_queue.get_jobs_by_name("auto_rapport"):
        job.schedule_removal()
    context.job_queue.run_daily(
        job_rapport_auto,
        time=datetime.time(hour=14, minute=0),  # 14h UTC = 15h Cameroun
        chat_id=chat_id,
        name="auto_rapport"
    )
    await update.message.reply_text(
        "✅ *Envoi automatique activé !*\n"
        "Vous recevrez votre rapport chaque jour à *15h00*.\n"
        "Pour arrêter : /stop",
        parse_mode="Markdown"
    )

async def cmd_stop(update: Update, context: ContextTypes.DEFAULT_TYPE):
    for job in context.job_queue.get_jobs_by_name("auto_rapport"):
        job.schedule_removal()
    await update.message.reply_text("🔕 Envoi automatique désactivé.")

async def bouton(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if q.data == "today":
        await _envoyer(q, datetime.date.today())
    elif q.data == "hier":
        await _envoyer(q, datetime.date.today() - datetime.timedelta(days=1))
    elif q.data == "auto_on":
        chat_id = q.message.chat_id
        for job in context.job_queue.get_jobs_by_name("auto_rapport"):
            job.schedule_removal()
        context.job_queue.run_daily(
            job_rapport_auto,
            time=datetime.time(hour=14, minute=0),
            chat_id=chat_id,
            name="auto_rapport"
        )
        await q.message.reply_text(
            "✅ Envoi automatique activé à *15h00* chaque jour !",
            parse_mode="Markdown"
        )
    elif q.data == "aide":
        await cmd_aide(update, context)

async def job_rapport_auto(context: ContextTypes.DEFAULT_TYPE):
    chat_id = context.job.chat_id
    date    = datetime.date.today()
    try:
        chemin, ivs, duree = generer_rapport(date)
        with open(chemin, "rb") as f:
            await context.bot.send_document(
                chat_id=chat_id,
                document=f,
                filename=os.path.basename(chemin),
                caption=f"🕒 *Rapport automatique 15h00 — {date.strftime('%d/%m/%Y')}*\n"
                        f"📊 {len(ivs)} interventions | ⏱ {duree}",
                parse_mode="Markdown"
            )
        os.remove(chemin)
    except Exception as e:
        await context.bot.send_message(chat_id=chat_id, text=f"❌ Erreur rapport auto : {e}")

# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────

def main():
    print("🤖 Bot IDCAM démarré...")
    app = ApplicationBuilder().token(TOKEN_BOT).build()
    app.add_handler(CommandHandler("start",   start))
    app.add_handler(CommandHandler("rapport", cmd_rapport))
    app.add_handler(CommandHandler("hier",    cmd_hier))
    app.add_handler(CommandHandler("aide",    cmd_aide))
    app.add_handler(CommandHandler("auto",    cmd_auto))
    app.add_handler(CommandHandler("stop",    cmd_stop))
    app.add_handler(CallbackQueryHandler(bouton))
    print("✅ En écoute...")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
