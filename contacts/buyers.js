// Buyer contacts — Anthropologie & Nuuly

// ─── Anthropologie ────────────────────────────────────────────────────────────
const anthropologie = {
  'BLOUSES & SHIRTS': [
    { name: 'Aly Kauffman',        email: 'AKauffman@anthropologie.com' },
    { name: 'Kate Creswell',       email: 'kcreswell@anthropologie.com' },
    { name: 'Danielle Coccerino',  email: 'dcoccerino@anthropologie.com' },
    { name: 'Madison Morley',      email: 'MMorley@anthropologie.com' },
  ],
  'DRESS & ROMPER': [
    { name: 'Lizzy Barrett',       email: 'ebarrett@anthropologie.com' },
    { name: 'Ellary Billings',     email: 'ebillings@anthropologie.com' },
    { name: 'Julie Stampone',      email: 'jstampone@anthropologie.com' },
    { name: 'Dami Amato',          email: 'damato@anthropologie.com' },
    { name: 'Molly MacRae',        email: 'mmacrae@anthropologie.com' },
  ],
  'PANTS & JUMPSUIT': [
    { name: 'Samantha Murphy',              email: 'smurphy@anthropologie.com' },
    { name: 'Caroline Riley',               email: 'criley4@anthropologie.com' },
    { name: "Gabriella Scotto D'Antuono",   email: 'gscottodantuono@anthropologie.com' },
  ],
  'LOUNGE': [
    { name: 'Simone Butler',       email: 'sbutler@anthropologie.com' },
    { name: 'Mackenzie Kroh',      email: 'mkroh@anthropologie.com' },
    { name: 'Olivia Schwann',      email: 'oschwann@anthropologie.com' },
    { name: 'Rozina Rum',          email: 'rrum@anthropologie.com' },
  ],
  'SKIRTS & SHORTS': [
    { name: 'Abby Podolsky',       email: 'apodolsky@anthropologie.com' },
    { name: 'Justin Leonard',      email: 'jleonard@anthropologie.com' },
    { name: 'Bella Kese',          email: 'bkese@anthropologie.com' },
  ],
  'SWIMWEAR': [
    { name: 'Ivy Turner',          email: 'iheilman@anthropologie.com' },
    { name: 'Vivian Nguyen',       email: 'vnguyen@anthropologie.com' },
    { name: 'Allisyn Blazier',     email: 'ablazier@anthropologie.com' },
    { name: 'Riley Longenderfer',  email: 'rlongenderfer@anthropologie.com' },
  ],
};

// ─── Nuuly ────────────────────────────────────────────────────────────────────
const nuuly = {
  'BLOUSES & SHIRTS': [
    { name: 'Emily Gallant',   email: 'egallant@urbn.com' },
    { name: 'Nuuly Tops',      email: 'nuulytops@urbn.com' },
  ],
  'DRESS, ROMPER & SWIMWEAR': [
    { name: 'Julia Dame',        email: 'jdame@urbn.com' },
    { name: 'Aaron Cooperman',   email: 'acooperman@urbn.com' },
    { name: 'Shelby Jensen',     email: 'sjensen1@urbn.com' },
    { name: 'Nuuly OnePieces',   email: 'nuulyonepieces@urbn.com' },
  ],
  'PANTS, JUMPSUIT, SKIRTS & SHORTS': [
    { name: 'Lydia Lepping',    email: 'llepping@urbn.com' },
    { name: 'Katelyn Buwalda',  email: 'kbuwalda@g.urbanout.com' },
    { name: 'Nuuly Bottoms',    email: 'NuulyBottoms@urbn.com' },
    { name: 'Hanna Saxon',      email: 'hsaxon@urbn.com' },
    { name: 'MC Miskuly',       email: 'mcmiskuly@urbn.com' },
  ],
};

// ─── Farm Rio internal (base PO emails) ──────────────────────────────────────
const anthroBasePO = [
  'anacarolina.azevedo@farmrio.com',
  'inbound@farmrio.com',
  'danielle.gouvea@farmrio.com',
];

const anthroBasePO_CC = [
  'paula@creativetwotwelve.com',
  'rafaela@showroom212.com',
  'ozan.guruscu@creativetwotwelve.com',
  'business@creativetwotwelve.com',
  'kamilla@creativetwotwelve.com',
];

module.exports = { anthropologie, nuuly, anthroBasePO, anthroBasePO_CC };
