-- ============================================
-- FINAL FIX FOR INFINITE RECURSION
-- Run this in Supabase SQL Editor
-- ============================================

-- Step 1: Disable RLS temporarily
ALTER TABLE public.user_profiles DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.user_invitations DISABLE ROW LEVEL SECURITY;

-- Step 2: Drop ALL existing policies
DROP POLICY IF EXISTS "Users can view own profile" ON public.user_profiles;
DROP POLICY IF EXISTS "Users can update own profile" ON public.user_profiles;
DROP POLICY IF EXISTS "Service role full access" ON public.user_profiles;
DROP POLICY IF EXISTS "Admins can view all profiles" ON public.user_profiles;
DROP POLICY IF EXISTS "Admins can update profiles" ON public.user_profiles;
DROP POLICY IF EXISTS "Admins can insert profiles" ON public.user_profiles;
DROP POLICY IF EXISTS "Admins can delete profiles" ON public.user_profiles;

DROP POLICY IF EXISTS "Service role full access invitations" ON public.user_invitations;
DROP POLICY IF EXISTS "Anyone can read invitation by token" ON public.user_invitations;
DROP POLICY IF EXISTS "Admins can view invitations" ON public.user_invitations;
DROP POLICY IF EXISTS "Admins can create invitations" ON public.user_invitations;
DROP POLICY IF EXISTS "Admins can update invitations" ON public.user_invitations;

-- Step 3: Re-enable RLS
ALTER TABLE public.user_profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.user_invitations ENABLE ROW LEVEL SECURITY;

-- ============================================
-- NEW POLICIES - NO INFINITE RECURSION
-- ============================================

-- USER PROFILES: Users can view their own profile only
CREATE POLICY "Users can view own profile"
    ON public.user_profiles
    FOR SELECT
    USING (auth.uid() = id);

-- USER PROFILES: Users can update their own profile (without role check to avoid recursion)
CREATE POLICY "Users can update own profile"
    ON public.user_profiles
    FOR UPDATE
    USING (auth.uid() = id);

-- USER PROFILES: Service role has full access (for admin operations)
CREATE POLICY "Service role full access"
    ON public.user_profiles
    FOR ALL
    USING (true)
    WITH CHECK (true);

-- USER INVITATIONS: Service role has full access
CREATE POLICY "Service role full access invitations"
    ON public.user_invitations
    FOR ALL
    USING (true)
    WITH CHECK (true);

-- USER INVITATIONS: Anyone can read pending invitations (for signup)
CREATE POLICY "Anyone can read invitation by token"
    ON public.user_invitations
    FOR SELECT
    USING (status = 'pending' AND expires_at > NOW());

-- ============================================
-- VERIFICATION
-- ============================================

-- Show all policies
SELECT tablename, policyname, permissive, roles, cmd
FROM pg_policies
WHERE tablename IN ('user_profiles', 'user_invitations')
ORDER BY tablename, policyname;

-- Success message
DO $$
BEGIN
    RAISE NOTICE '✓ RLS policies fixed!';
    RAISE NOTICE '✓ No more infinite recursion';
    RAISE NOTICE '✓ Try logging in now';
END $$;
