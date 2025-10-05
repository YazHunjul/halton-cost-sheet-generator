-- ============================================
-- Cleanup Cancelled/Expired Invitations
-- Run this to delete old invitations that are blocking new ones
-- ============================================

-- Delete all cancelled invitations
DELETE FROM public.user_invitations
WHERE status = 'cancelled';

-- Delete all expired invitations
DELETE FROM public.user_invitations
WHERE status = 'expired';

-- Optional: Delete all accepted invitations (users already created)
DELETE FROM public.user_invitations
WHERE status = 'accepted';

-- Show remaining invitations
SELECT
    email,
    first_name,
    last_name,
    role,
    status,
    created_at,
    expires_at
FROM public.user_invitations
ORDER BY created_at DESC;

-- Success message
DO $$
BEGIN
    RAISE NOTICE '✓ Cancelled and expired invitations deleted';
    RAISE NOTICE '✓ You can now create new invitations with the same emails';
END $$;
