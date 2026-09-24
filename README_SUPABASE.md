# Épargne Étudiant — Supabase

## Fichiers

- `epargne.py` : application Streamlit
- `requirements.txt` : dépendances Python
- `.streamlit/secrets.toml.example` : modèle des secrets
- `supabase_schema.sql` : schéma PostgreSQL
- `migrate_sqlite_to_supabase.py` : migration de l'ancienne base SQLite
- `.gitignore` : protection des secrets et bases locales

## 1. Installer

```bash
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
```

Linux/macOS:

```bash
source .venv/bin/activate
```

Puis:

```bash
pip install -r requirements.txt
```

## 2. Créer les tables Supabase

Dans Supabase:

`SQL Editor` → nouveau script → coller `supabase_schema.sql` → Run.

## 3. Configurer les secrets

Copier:

```text
.streamlit/secrets.toml.example
```

vers:

```text
.streamlit/secrets.toml
```

Puis mettre le nouveau mot de passe Supabase.

NE PAS publier `secrets.toml`.

## 4. Migrer l'ancienne base

Placez `epargne_etudiant.db` à côté de `migrate_sqlite_to_supabase.py`.

Définissez `SUPABASE_PASSWORD`, puis:

```bash
python migrate_sqlite_to_supabase.py
```

Le script copie:
- membres
- cotisations
- emprunts
- échéances et paiements

Il conserve les identifiants relationnels via une table de correspondance interne pendant la migration.

## 5. Lancer

```bash
streamlit run epargne.py
```

## 6. WhatsApp

Numéro administrateur fixe de l'application:

`+221 77 752 19 69`

Les liens WhatsApp préremplis peuvent fonctionner sans Twilio.

Pour l'envoi automatique, renseignez les variables Twilio dans `secrets.toml`.

## Sécurité

Le mot de passe Supabase qui a été partagé précédemment doit être considéré comme compromis.
Régénérez-le dans Supabase et utilisez uniquement le nouveau mot de passe dans `secrets.toml`.

Pour une version production, il faudra aussi remplacer le stockage des mots de passe administrateurs en clair par des mots de passe hachés.
