-- ============================================
-- Companies Management Schema
-- Add this to your existing database
-- ============================================

-- ============================================
-- COMPANIES TABLE
-- ============================================
CREATE TABLE IF NOT EXISTS public.companies (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    name TEXT UNIQUE NOT NULL,
    address TEXT NOT NULL,
    is_active BOOLEAN DEFAULT true,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    created_by UUID REFERENCES auth.users(id),
    updated_by UUID REFERENCES auth.users(id)
);

-- Index for faster queries
CREATE INDEX IF NOT EXISTS idx_companies_name ON public.companies(name);
CREATE INDEX IF NOT EXISTS idx_companies_active ON public.companies(is_active);

-- ============================================
-- RLS POLICIES FOR COMPANIES
-- ============================================

-- Enable RLS
ALTER TABLE public.companies ENABLE ROW LEVEL SECURITY;

-- Everyone can view active companies (for dropdowns)
CREATE POLICY "Anyone can view active companies"
    ON public.companies
    FOR SELECT
    USING (is_active = true);

-- Service role can do everything (for admin operations)
CREATE POLICY "Service role full access companies"
    ON public.companies
    FOR ALL
    USING (current_user = 'service_role')
    WITH CHECK (current_user = 'service_role');

-- ============================================
-- TRIGGER FOR UPDATED_AT
-- ============================================

DROP TRIGGER IF EXISTS update_companies_updated_at ON public.companies;
CREATE TRIGGER update_companies_updated_at
    BEFORE UPDATE ON public.companies
    FOR EACH ROW
    EXECUTE FUNCTION public.update_updated_at();

-- ============================================
-- IMPORT EXISTING COMPANIES
-- ============================================
-- This imports all existing companies from your business_data.py

INSERT INTO public.companies (name, address, is_active) VALUES
('Airedale Group (Bradford)', E'Victoria Road\nPedeshill\nBradford BD23 2BN', true),
('Airedale Group (Lutterworth)', E'1 St Johns Business Park\nRugby Road\nLutterworth LE17 4HB', true),
('Court Catering Equipment Ltd', E'Unit 1, Acton Vale Ind. Park,\nCowley Road\nLondon W3 7XA', true),
('Humble Arnold Associates', E'Farriers House\nFarriers Close\nCodicote\nHertfordshire\nSG4 8DU', true),
('Berkeley Projects UK Ltd', E'17 Ewell Road\nCheam\nSurrey\nSM3 8DD', true),
('Chapman Ventilation Ltd', E'15-20 Woodfield Rd\nWelwyn Garden City\nHerts AL7 1JQ', true),
('SG Group', E'Aspen Way\nPaignton\nDevon TQ4 7QR', true),
('ABM Catering for Leisure Ltd', E'Algate House\nClydesmuir Rd Ind Est\nCardiff CF24 2QS', true),
('C&C Catering Equipment Ltd', E'1 Smithy Farm\nChapel Lane\nSaighton\nChester CH3 6EW', true),
('Hallmark Kitchens Ltd', E'South Barn\nCrockham Farm\nEdenbridge\nKent\nTN8 6SR', true),
('Cabiola Foodservice Equipment', E'The Bake House\nNarborough Wood Park\nDesford Road\nEnderby\nLeicester LE19 4XT', true),
('Ceba Catering Services', E'Unit 27, Eastville Close\nEastern Avenue Trading Estate\nGloucester\nGL4 3SJ', true),
('Design Installation Service Ltd', E'4 Gainsborough House\n42 / 44 Bath Road\nCheltenham\nGL53 7HJ', true),
('Nelson Bespoke Commercial Kitchens', E'Unit 1\nRowley Industrial Park\nRoslin Road\nActon\nLondon W3 8BH', true),
('Airflow Cooling Ltd', E'132 Rutland Road\nSheffield S3 9PP', true),
('Spectrum Contracts Ltd.', E'Unit 11 Dorcan Business Village\nMurdoch Road\nDorcan\nSwindon\nSN3 5HY', true),
('CCE Group Ltd', E'Unit 1 Bentley Farm\nOld Church Hill\nChingon Hills\nBasildon SS16 6HZ', true),
('Shine Food Machinery Ltd', E'New Quay Road\nStephenson Street Ind. Est.\nNewport NP19 4FL', true),
('HCE Foodservice Equipment Ltd', E'School Lane\nChandlers Ford\nEastleigh\nHants SO53 4DG', true),
('Vision Commercial Kitchens Limited', E'Unit A1, Axis Point,\nHill Top Road,\nHeywood\nLancs OL10 2RQ', true),
('Fast Food Systems Ltd.', E'Unit 1\nHeadley Park 9\nHeadley Road East\nWoodley\nReading\nBerkshire RG5 4SQ', true),
('C. Caswell (Eng) Services Ltd', E'Knowsley Road Ind. Est.\nHaslingden\nRossendale\nLancs BB4 4RX', true),
('VSS Ltd', E'Building 2 ThermoAir Site\nAthy Road\nCarlow\nIreland R93 K635', true),
('Reco-Air', E'Newmarket 24, Centrix,\nKeys Business Village\nCannock\nStaffordshire\nWS12 2HA', true),
('Gratte Brothers Ltd.', E'3 Crompton Road\nStevenage\nHertfordshire\nSG1 2XP', true),
('Salix Stainless Steel Production House', E'Chester Hall Lane\nBasildon\nEssex SS14 3BG', true),
('Stangard Design Solutions Ltd.', E'The Dairy\nWolvey Lodge Business Centre\nWolvey\nWarwickshire\nLE10 3HB', true),
('AFR Air Conditioning Ltd', E'14 Kingsdown Road\nSwindon\nWiltshire\nSN25 6PB', true),
('Sigma Catering Equipment', E'Unit 4\nPotter Street\nWallsend\nNE28 6LS', true),
('CHR Equipment Ltd.', E'Astar House\nFourries Bridge\nPreston\nPR5 6GS', true),
('MITIE Engineering Services (Retail) Ltd.', E'The Millennium Centre\nM4 Crosby Way\nFarnham\nSurrey GU9 7XX', true),
('Catering Design Services', E'4 Waterside Commerce Park\nTrafford Park\nManchester\nM17 1WD', true),
('Salix Stainless Steel Applications', E'Production House\nChester Hall Lane\nBasildon\nEssex SS14 3BG', true),
('Galgorm Group', E'Galgorm Industrial Estate\n7 Corbally Road\nBallymena BT42 1JQ', true),
('Michael J Lonsdale', E'Unit 1Langley Quay\nWaterside Drive\nLangley\nSlough\nBerkshire\nSL3 6EY', true),
('Advance Catering Equipment', E'Dantem House\nBlackburn Road\nHoughton Regis\nDunstable\nLU5 5BQ', true),
('MITIE Engineering Services (South West) Ltd', E'5 Hanover Court\nMatford Close\nMatford Business Park\nExeter\nEX2 8QJ', true),
('A C V R Services Ltd', E'Unit 4\nNewtown Grange Farm Bu siness Park Orchard\nLeicester\nLE8 8FL', true),
('Lockhart Catering Equipment', E'6 The Astra Centre\nEdinburgh Way\nHarlow\nEssex\nCM20 2BN', true),
('Insitu Fabrications', E'Unit19\nButterfield Ind. Est.\nNewtongrange\nBonnyrigg EH19 3JQ', true),
('Stephens Catering Equipment', E'205 Carnabanaugh Road\nDoughandshane\nBallymena\nCo. Antrim BT42 4NY', true),
('Into Design Ltd', E'16 Richards Rise\nOxted\nSurrey\nRH8 0TS', true),
('Merrick Contracts Ltd', E'Bird Point Mill\nKing Henry''s Drive\nNew Addington\nSurrey\nCR0 0AE', true),
('CKE Service Ltd', E'13 High Street\nThames Ditton\nSurrey\nKT7 0SN', true),
('Dentons Catering Equipment', E'2/4 Clapham High Street\nLondon\nSW4 7UT', true),
('Ferro Design Ltd', E'Blay''s House\nChurchfield Road\nChalfont St Peter\nBucks\nSL9 9EW', true),
('Four Seasons Air Conditioning Suppliers Ltd', E'Stadium Works\nSedgley Street\nWolverhampton\nWest Midlands\nWV2 3AJ', true),
('Pan-Euro Environmental', E'Unit 2, Group House\nAlbon Way\nReigate\nRH2 7JY', true),
('Caswell Engineering', E'Knowsley Road Ind. Est.\nHaslingden\nRossendale\nLancs BB4 4RR', true),
('C A Sothers Ltd', E'156 Hockey Hill\nBirmingham\nB18 5AN', true),
('RDA', E'5 Apollo Court\nMonkton Business Park South\nTyne & Wear NE31 2ES', true),
('KCCJ Ltd.', E'The Old Granary\nCourt Lodge Farm\nLambden Hill\nDA2 7QY', true),
('Contract Catering Equipment', E'Unit1, Bentley Farm,\nOld Church Hill, Chingon Hills,\nEssex SS16 6HZ', true),
('Fellerman Partnership Ltd', E'74 Kimberley Road\nPortsmouth PO4 9NS', true),
('Tricon Foodservice Consultants', E'St James House\n27-43 Eastern Road\nRomford', true),
('Technical Services Ref & A/C Ltd', E'Arlingston House\n32 Boundary Road\nNewbury', true),
('Sefton Horn Winch', E'The Stables\nHome Farm Business Units\nRiverside\nEynsford', true),
('Carford Group', E'1-4 Mitchell Road\nFernside Park\nFerndown Industrial Estate\nFerndown\nDorset BH21 7SG', true),
('Gratte Bros. Catering Equipment Ltd', E'3 Crompton Road\nStevenage\nHerts SG1 2XP', true),
('Scobie & McIntosh', E'15 Brewster Square\nBrucefield Industry Park\nLivingston\nWest Lothian\nEH54 9BJ', true),
('SVS Ltd', E'Unit 3\nGreencroft Ind. Est.\nAnnfield Plain\nStanley\nCo Durham\nDH9 7YB', true),
('Kitchequip', E'Canal View\nWaterside Business Park\nNew Lane\nBurscough L40 8JX', true),
('Space Catering', E'Barnwood Point\nCorinium Avenue\nGlos GL4 3HX', true),
('GS Catering Equipment Ltd', E'Aspen Way\nPaignton\nDevon TQ4 7QR', true),
('GastroNorth Ltd', E'Merlin House\nTeam Valley Trading Est\nPrinces Park\nGateshead\nNE11 0NF', true),
('Promart Manufacturing Ltd', E'2B Caddick Road\nKnowsley Business Park\nPrescot\nMerseyside L34 9HP', true),
('Bob Sackett Commercial Catering Projects Ltd', E'9 Chesterford Court\nLondon Road\nGreat Chesterford\nEssex CB10 1PF', true),
('Main Contract Services Ltd', E'Alexandra House\nStation Road\nGrangemouth FK3 8DG', true),
('CDS-Wilman', E'4 Waterside Commerce Park\nTrafford Park\nManchester M17 1W', true),
('Western Blueprint Ltd', E'Unit B2, 1st Floor\nTrym House Business Centre\nForest Road\nKingswood\nBristol BS15 8DH', true),
('Edge Design', E'Unit 3 Wiston House\nWiston Avenue\nWorthing\nWest Sussex BN14 7QL', true),
('Chapman Ventilation', E'15 - 20 Woodfield Road\nWelwyn Garden City\nHertfordshire\nAL7 1JQ', true),
('RecoAir', E'14 Heritage Park\nHayes Way\nCannock\nStaffordshire WS11 7LT', true),
('Humble Arnold', E'Farriers House,\nFarriers Close,\nCodicote\nHerts SG4 8DU', true)
ON CONFLICT (name) DO NOTHING;

-- Success message
DO $$
BEGIN
    RAISE NOTICE 'Companies table created and populated successfully!';
END $$;
