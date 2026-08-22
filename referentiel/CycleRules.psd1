@{
    # Ordre = ordre exact de test dans E5Controller::detectCycle (premier prefixe correspondant gagne ; sinon 'autres')
    Rules = @(
        @{ Order = 1;  Prefix = '606880'; Cycle = 'emballages' }
        @{ Order = 2;  Prefix = '6068';   Cycle = 'emballages' }
        @{ Order = 3;  Prefix = '606';    Cycle = 'energie_dechets' }
        @{ Order = 4;  Prefix = '164';    Cycle = 'tresorerie_financement' }
        @{ Order = 5;  Prefix = '20';     Cycle = 'immobilisations' }
        @{ Order = 6;  Prefix = '21';     Cycle = 'immobilisations' }
        @{ Order = 7;  Prefix = '22';     Cycle = 'immobilisations' }
        @{ Order = 8;  Prefix = '23';     Cycle = 'immobilisations' }
        @{ Order = 9;  Prefix = '28';     Cycle = 'immobilisations' }
        @{ Order = 10; Prefix = '40';     Cycle = 'fournisseurs' }
        @{ Order = 11; Prefix = '41';     Cycle = 'clients' }
        @{ Order = 12; Prefix = '32';     Cycle = 'stocks' }
        @{ Order = 13; Prefix = '60';     Cycle = 'charges' }
        @{ Order = 14; Prefix = '61';     Cycle = 'services' }
        @{ Order = 15; Prefix = '62';     Cycle = 'personnel' }
        @{ Order = 16; Prefix = '63';     Cycle = 'fiscal' }
        @{ Order = 17; Prefix = '64';     Cycle = 'financier' }
    )
    DefaultCycle = 'autres'
}
