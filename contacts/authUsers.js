// Workspace login users — passwords are env vars (override via WORKSPACE AUTH Google Sheet tab)
// If env var not set, the default password applies — user is forced to change on first login.
const DEFAULT_PASSWORD = '12345';

module.exports = {
  DEFAULT_PASSWORD,
  users: [
    { email: 'support@creativetwotwelve.com',       envVar: 'ADMIN_PASSWORD',    name: 'Flavio Azevedo',   role: 'admin' },
    { email: 'business@creativetwotwelve.com',       envVar: 'MANUELA_PASSWORD',  name: 'Manuela Carvalho', role: 'user'  },
    { email: 'samples@creativetwotwelve.com',        envVar: 'IGO_PASSWORD',      name: 'Igo Gardel',       role: 'user'  },
    { email: 'logistics@creativetwotwelve.com',      envVar: 'EDUARDO_PASSWORD',  name: 'Eduardo Moraes',   role: 'user'  },
    { email: 'inspection@creativetwotwelve.com',     envVar: 'JULIAN_PASSWORD',   name: 'Julian Fajardo',   role: 'user'  },
    { email: 'paula@creativetwotwelve.com',          envVar: 'PAULA_PASSWORD',    name: 'Paula Erthal',     role: 'user'  },
    { email: 'rafaela.neves@farmrio.com',            envVar: 'RAFAELA_PASSWORD',  name: 'Rafaela Neves',    role: 'user'  },
    { email: 'ozan.guruscu@creativetwotwelve.com',   envVar: 'OZAN_PASSWORD',     name: 'Ozan Guruscu',     role: 'user'  },
    { email: 'kamilla@creativetwotwelve.com',        envVar: 'KAMILLA_PASSWORD',  name: 'Kamilla Aguiar',   role: 'user'  },
  ],
};
