@{
    Missions = @(
        @{
            Code = 'ENERGY'; Name = 'Diagnostic énergétique'
            Description = 'Analyse des consommations et préconisations d''économies d''énergie'
            BasePrice = 2500
            Services = @('Audit énergétique complet', 'Mise en place d''indicateurs de suivi', 'Élaboration d''un plan d''action', 'Reporting trimestriel des consommations')
            Tools = @('Outil énergie 1', 'Outil énergie 2')
        }
        @{
            Code = 'REGUL'; Name = 'Mise en conformité réglementaire'
            Description = 'Accompagnement pour la conformité aux réglementations environnementales'
            BasePrice = 1800
            Services = @('Diagnostic de conformité', 'Analyse des écarts', 'Plan de mise en conformité', 'Formation des équipes')
            Tools = @('Kit ICPE', 'Guide des obligations RSE')
        }
        @{
            Code = 'CARBON'; Name = 'Bilan Carbone®'
            Description = 'Réalisation ou évaluation d''un bilan carbone selon la méthode ABC'
            BasePrice = 3500
            Services = @('Calcul des émissions', 'Rapport détaillé', 'Plan de réduction', 'Obtention du label')
            Tools = @('Kit Bilan Carbone®', 'Plateforme ABC')
        }
        @{
            Code = 'SUPPLY'; Name = 'Analyse de la chaîne d''approvisionnement'
            Description = 'Évaluation des risques et opportunités dans la supply chain'
            BasePrice = 2200
            Services = @('Cartographie des fournisseurs', 'Analyse des risques', 'Plan d''amélioration', 'Formation achats responsables')
            Tools = @('Outil fournisseurs', 'Grille d''évaluation RSE')
        }
        @{
            Code = 'REPORT'; Name = 'Reporting RSE'
            Description = 'Accompagnement à la rédaction d''un rapport de durabilité'
            BasePrice = 4000
            Services = @('Collecte des données', 'Analyse des écarts CSRD', 'Rédaction du rapport', 'Mise en forme')
            Tools = @('Kit CSRD', 'Plateforme Impact Durabilité')
        }
    )
}
