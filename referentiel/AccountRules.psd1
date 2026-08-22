@{
    # Ordre = ordre exact de test dans E5Controller::generateAnalysisForLine (premier prefixe correspondant gagne)
    Rules = @(
        @{ Order = 1;  Prefix = '215';    Reading = 'Investissement en matériel industriel - vérifier l''efficience énergétique'; Question = 'Avez-vous considéré des équipements à meilleure performance énergétique?'; Suggestion = 'Étudier les aides pour l''achat d''équipements verts' }
        @{ Order = 2;  Prefix = '2183';   Reading = 'Investissement en matériel de transport - opportunité pour véhicules propres'; Question = 'Pourriez-vous intégrer des véhicules électriques ou hybrides?'; Suggestion = 'Bénéficiez du bonus écologique pour les véhicules propres' }
        @{ Order = 3;  Prefix = '401';    Reading = 'Relations avec les fournisseurs - opportunité d''achats responsables'; Question = 'Avez-vous mis en place une politique d''achats responsables?'; Suggestion = 'Évaluer l''empreinte environnementale des principaux fournisseurs' }
        @{ Order = 4;  Prefix = '411';    Reading = 'Relations clients - sensibilité aux critères RSE'; Question = 'Vos clients vous sollicitent-ils sur des critères environnementaux?'; Suggestion = 'Développer des arguments écologiques pour vos produits/services' }
        @{ Order = 5;  Prefix = '322';    Reading = 'Gestion des stocks - risque de gaspillage ou surstockage'; Question = 'Constatez-vous un niveau de surstockage ou de pertes?'; Suggestion = 'Mettre en place des solutions de valorisation des invendus' }
        @{ Order = 6;  Prefix = '3220';   Reading = 'Stocks de matières premières - impact environnemental'; Question = 'Avez-vous évalué l''empreinte écologique de vos matières premières ?'; Suggestion = 'Étudier des alternatives plus durables' }
        @{ Order = 7;  Prefix = '32200';  Reading = 'Stocks de fournitures - gestion durable'; Question = 'Pourriez-vous réduire les emballages à usage unique ?'; Suggestion = 'Mettre en place un système de consigne' }
        @{ Order = 8;  Prefix = '625';    Reading = 'Frais de déplacement - impact environnemental'; Question = 'Avez-vous une politique de gestion des déplacements?'; Suggestion = 'Envisager des mobilités alternatives (véhicules hybrides, télétravail)' }
        @{ Order = 9;  Prefix = '606';    Reading = 'Consommations énergétiques - potentiel d''économie'; Question = 'Disposez-vous de données sur vos consommations d''énergie?'; Suggestion = 'Mettre en place des actions d''économie d''énergie' }
        @{ Order = 10; Prefix = '164';    Reading = 'Financement - opportunité de solutions vertes'; Question = 'Avez-vous bénéficié de prêts ou subventions liés à l''écologie?'; Suggestion = 'Explorer les solutions de financement vert (emprunts à impact, subventions)' }
        @{ Order = 11; Prefix = '606880'; Reading = 'Consommation d''emballages - potentiel de réduction'; Question = 'Avez-vous envisagé des solutions d''emballages réutilisables?'; Suggestion = 'Mettre en place une politique d''emballages durables' }
    )
}
