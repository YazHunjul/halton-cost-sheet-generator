-- ============================================
-- FIXED RLS POLICIES - Run this to fix infinite recursion
-- ============================================

-- First, drop all existing policies on user_profiles
DROP POLICY IF EXISTS "Admins can view all profiles" ON public.user_profiles;
DROP POLICY IF EXISTS "Users can view own profile" ON public.user_profiles;
DROP POLICY IF EXISTS "Admins can update profiles" ON public.user_profiles;
DROP POLICY IF EXISTS "Users can update own profile" ON public.user_profiles;
DROP POLICY IF EXISTS "Admins can insert profiles" ON public.user_profiles;
DROP POLICY IF EXISTS "Admins can delete profiles" ON public.user_profiles;

-- Drop policies on invitations
DROP POLICY IF EXISTS "Admins can view invitations" ON public.user_invitations;
DROP POLICY IF EXISTS "Admins can create invitations" ON public.user_invitations;
DROP POLICY IF EXISTS "Admins can update invitations" ON public.user_invitations;

-- ============================================
-- FIXED USER PROFILES POLICIES (No infinite recursion)
-- ============================================

-- Users can view their own profile
CREATE POLICY "Users can view own profile"
    ON public.user_profiles
    FOR SELECT
    USING (auth.uid() = id);

-- Users can update their own profile (but not their role)
CREATE POLICY "Users can update own profile"
    ON public.user_profiles
    FOR UPDATE
    USING (auth.uid() = id)
    WITH CHECK (
        auth.uid() = id
        -- Prevent users from changing their own role
        AND role = (SELECT role FROM public.user_profiles WHERE id = auth.uid())
    );

-- Service role can do everything (bypasses RLS)
-- This allows admin operations through service_role key
CREATE POLICY "Service role full access"
    ON public.user_profiles
    FOR ALL
    USING (current_setting('request.jwt.claims', true)::json->>'role' = 'service_role')
    WITH CHECK (current_setting('request.jwt.claims', true)::json->>'role' = 'service_role');

-- ============================================
-- FIXED USER INVITATIONS POLICIES
-- ============================================

-- Service role can do everything with invitations
CREATE POLICY "Service role full access invitations"
    ON public.user_invitations
    FOR ALL
    USING (current_setting('request.jwt.claims', true)::json->>'role' = 'service_role')
    WITH CHECK (current_setting('request.jwt.claims', true)::json->>'role' = 'service_role');

-- Allow anyone to read pending invitations by token (for registration)
CREATE POLICY "Anyone can read invitation by token"
    ON public.user_invitations
    FOR SELECT
    USING (status = 'pending' AND expires_at > NOW());

-- ============================================
-- VERIFICATION
-- ============================================

-- List all policies to verify
SELECT schemaname, tablename, policyname, permissive, roles, cmd, qual
FROM pg_policies
WHERE tablename IN ('user_profiles', 'user_invitations')
ORDER BY tablename, policyname;
