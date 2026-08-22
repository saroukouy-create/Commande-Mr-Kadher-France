@{
    Profiles = @(
        @{
            Key = 'Proactif'; MinScore = 24; MaxScore = 30; Label = 'Profil 4 : PROACTIF'; Header = 'Félicitations !'
            Positionnement = 'Déjà engagé dans une stratégie claire. En recherche d''optimisation, de pilotage ou de valorisation plus poussée.'
            Freins = @(
                'Complexité de mise en œuvre'
                'Isolement ou manque d''outils de pilotage'
                'Besoin de reconnaissance externe'
            )
            Arguments = @(
                'Nous pouvons intégrer la dimension écologique dans vos tableaux de bord et vous aider à produire votre futur rapport de durabilité.'
                'Vous êtes éligible à de nouveaux financements verts : nous pouvons vous aider à les obtenir.'
                'Nous vous proposons des indicateurs personnalisés pour suivre l''impact de vos actions.'
            )
            Ton = 'expert, valorisant l''excellence, la performance, la conformité et l''anticipation réglementaire.'
        }
        @{
            Key = 'Ouvert'; MinScore = 16; MaxScore = 23; Label = 'Profil 3 : OUVERT'; Header = 'Bon Potentiel !'
            Positionnement = 'Déjà sensibilisé, prêt à agir ou partiellement engagé. À la recherche d''un cadre ou d''un accompagnement structurant.'
            Freins = @(
                'Manque de méthode ou de stratégie claire'
                'Manque de ressources internes'
                'Besoin de valorisation externe'
            )
            Arguments = @(
                'Nous vous aidons à structurer votre démarche pour en faire un levier stratégique.'
                'L''analyse de vos comptes révèle des leviers de performance écologique.'
                'Votre engagement peut être mis en avant auprès de vos partenaires (label, communication, image de marque).'
            )
            Ton = 'mobilisateur, orienté sur la valeur stratégique et la montée en compétence.'
        }
        @{
            Key = 'Prudent / Intermédiaire'; MinScore = 8; MaxScore = 15; Label = 'Profil 2 : PRUDENT / INTERMÉDIAIRE'; Header = 'À Approfondir'
            Positionnement = 'Sensibilisé, mais hésitant à s''engager. Voit l''intérêt mais doute du retour sur investissement ou de la faisabilité.'
            Freins = @(
                'Peur d''un investissement trop lourd'
                'Doutes sur les compétences internes'
                'Incertitude sur les bénéfices concrets'
            )
            Arguments = @(
                'Nous pouvons commencer par une analyse simple de vos charges pour identifier les leviers d''action à court terme.'
                'L''accompagnement du cabinet vous évite de tout faire seul : outils, plan d''action, indicateurs, tout est clé en main.'
                'Vous pouvez avancer à votre rythme : la démarche est progressive et adaptée à votre structure.'
            )
            Ton = 'engageant, valorisant la souplesse et l''assistance, orienté solution simple et accompagnée.'
        }
        @{
            Key = 'Réticent'; MinScore = 0; MaxScore = 7; Label = 'Profil 1 : RÉTICENT'; Header = 'Recommandations'
            Positionnement = 'Ne voit aucun intérêt à la transition écologique, ni à court ni à long terme. Peu ou pas engagé, souvent dans le déni ou préoccupé uniquement par la rentabilité immédiate.'
            Freins = @(
                'Vision de la transition comme une contrainte réglementaire inutile ou coûteuse'
                'Absence de perception d''intérêt ou de bénéfices'
                'Manque de temps ou surcharge'
            )
            Arguments = @(
                'Vous pouvez réduire vos coûts d''énergie et vos frais de fonctionnement (ex. déchets, transport) sans bouleverser votre activité.'
                'De nouvelles aides financières ou incitations fiscales sont accessibles aux entreprises qui s''engagent, même modestement.'
                'De plus en plus de vos clients vous demanderont des preuves d''engagement écologique : anticiper ces attentes vous évite des ruptures de contrats.'
            )
            Ton = 'rassurant, pédagogique, orienté pragmatisme et gains économiques.'
        }
    )
}
