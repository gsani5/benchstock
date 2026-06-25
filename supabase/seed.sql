-- ============================================================
-- Optional starter inventory. Run AFTER schema.sql if you want
-- the lab's real reagents preloaded. Safe to skip or edit.
-- ============================================================

insert into items (name, category, lot_number, catalog_number, supplier, quantity, unit, reorder_threshold, location, expiration_date, notes) values
('Anti-mouse TNF-α (clone XT3.11)', 'Antibody', 'BP0058-7745', 'BP0058', 'Bio X Cell', 4, 'vials', 2, '−20 °C, Freezer B, Shelf 2', '2026-11-30', 'InVivoPlus. 600 µg i.p. dosing for TNFα blockade experiments.'),
('FITC-dextran (4 kDa)', 'Reagent', 'FD4-MKCQ2210', 'FD4', 'Sigma-Aldrich', 1, 'g', 1, '4 °C, Fridge A, Door', '2026-08-15', 'BBB permeability assay. Protect from light.'),
('Isolectin GS-IB4 — AF647', 'Reagent', 'I32450-2391', 'I32450', 'Thermo Fisher', 2, 'vials', 1, '−20 °C, Freezer B, Shelf 1', '2027-02-28', 'Vascular labeling for CLEM/EM.'),
('Paraformaldehyde 16% (EM grade)', 'Chemical', '15710-PFA-0042', '15710', 'EM Sciences', 6, 'ampoules', 8, 'Flammables cabinet, Bay 3', '2026-09-01', 'Perfusion fixative. P11 protocol.'),
('V-PLEX Proinflammatory Panel 1 (mouse)', 'Kit', 'K15048D-0188', 'K15048D', 'Meso Scale Discovery', 1, 'kit', 1, '−20 °C, Freezer C', '2026-07-20', 'Cytokine MSD on cerebellar + plasma samples.'),
('Cryomold (standard, 25×20×5 mm)', 'Consumable', '4557-LOT19', '4557', 'Sakura', 12, 'boxes', 4, 'Bench 4, Drawer 2', null, '');
