# Streamlit Cloud Deployment Checklist

Quick checklist for deploying to Streamlit Cloud.

## Pre-Deployment (Local Setup)

- [x] Updated `.env` with new Supabase credentials
- [x] Created `.streamlit/secrets.toml` for local testing
- [x] Updated `supabase_config.py` to support both .env and secrets.toml
- [x] Created admin user (Yazanhunjul5@gmail.com)
- [ ] Run `database/rls_fix_final.sql` in Supabase SQL Editor
- [ ] Run `database/companies_schema.sql` in Supabase SQL Editor
- [ ] Create Supabase Storage bucket named `templates` (set to Private)
- [ ] Test login locally: `streamlit run src/app_with_auth.py`

## Supabase Setup

- [ ] Database tables created (user_profiles, user_invitations, companies, audit_logs)
- [ ] RLS policies fixed (no infinite recursion)
- [ ] Storage bucket `templates` created
- [ ] Storage policies added (authenticated read, service_role full access)
- [ ] Admin user created and confirmed working
- [ ] Companies imported (should see 75+ companies)

## GitHub Repository

- [ ] All code pushed to GitHub
- [ ] `.env` is NOT committed (check `.gitignore`)
- [ ] `.streamlit/secrets.toml` is NOT committed (check `.gitignore`)
- [ ] `requirements.txt` includes all dependencies

## Streamlit Cloud Configuration

- [ ] Signed in to https://share.streamlit.io/
- [ ] Created new app
- [ ] Selected correct repository and branch
- [ ] Set main file path: `src/app_with_auth.py`
- [ ] Added secrets in Advanced Settings (copy from `.streamlit/secrets.toml`)
- [ ] Deployed app

## Post-Deployment Testing

- [ ] App deployed successfully
- [ ] Can access app URL: https://[your-app-name].streamlit.app
- [ ] Login works with admin credentials
- [ ] Admin Panel accessible
- [ ] User Management shows admin user
- [ ] Company Management shows all companies
- [ ] Template Management accessible
- [ ] Can upload templates via "First-Time Setup"
- [ ] Can create a test quotation
- [ ] Word document generates successfully
- [ ] Can download generated quotation

## Security Verification

- [ ] Changed admin password from default
- [ ] Service role key NOT exposed in code
- [ ] RLS enabled on all tables
- [ ] Storage bucket set to Private
- [ ] `.env` and `secrets.toml` in `.gitignore`

## Optional Enhancements

- [ ] Set up custom domain
- [ ] Configure email notifications (Supabase Auth settings)
- [ ] Add monitoring/logging
- [ ] Create additional admin users
- [ ] Customize app URL
- [ ] Set up backups (Supabase Dashboard)

## Troubleshooting Steps (If Needed)

If app doesn't work:

1. Check Streamlit Cloud logs for errors
2. Verify secrets are correctly added in Streamlit Cloud
3. Confirm Supabase RLS policies are fixed
4. Test Supabase connection from app
5. Check browser console for JavaScript errors
6. Verify all database tables exist
7. Confirm storage bucket and policies are set up

## Quick Reference

**Admin Login:**
- Email: Yazanhunjul5@gmail.com
- Password: Admin123!@# (CHANGE THIS!)

**Supabase Project:**
- URL: https://vxvjncrolwyvvykalirw.supabase.co
- Dashboard: https://supabase.com/dashboard/project/vxvjncrolwyvvykalirw

**Key Files:**
- App entry point: `src/app_with_auth.py`
- Supabase config: `src/config/supabase_config.py`
- Local secrets: `.streamlit/secrets.toml`
- Environment vars: `.env`

**Important SQL Files:**
- RLS fix: `database/rls_fix_final.sql`
- Companies: `database/companies_schema.sql`
- Complete schema: `database/complete_schema.sql`

## Notes

- Streamlit Cloud filesystem is ephemeral (files don't persist)
- Templates MUST be in Supabase Storage
- Use "First-Time Setup" in Admin Panel after deployment
- App auto-redeploys when you push to GitHub
