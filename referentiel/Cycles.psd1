@{
    Cycles = @(
        @{
            Key = 'immobilisations'; Name = 'Immobilisation'; Description = 'Analyse des investissements et équipements'
            Questions = @(
                'Avez-vous identifié des équipements obsolètes ou énergivores ?'
                'Certains matériels sont-ils concernés par des normes environnementales spécifiques ?'
                'Des réflexions sont-elles en cours sur la rénovation énergétique des bâtiments ?'
                'Avez-vous envisagé des aides pour financer des investissements "verts" ?'
            )
            Issues = @(
                'Équipements anciens/énergivores'
                'Bâtiment mal isolé ou non rénové'
                'Absence de suivi des consommations énergétiques'
                'Sortie d''immobilisation : la question du recyclage des sorties et des déchets dangereux'
                'Entrée d''immobilisation : non analyse d''impact écologique avant le choix et l''achat'
            )
            Actions = @(
                'Audit énergétique avec partenaire certifié'
                'Aide aux financements ADEME/Bpifrance'
                'Plan de rénovation thermique'
                'Programme de remplacement progressif des équipements'
                'Investissements dans des immobilisations durables'
            )
        }
        @{
            Key = 'fournisseurs'; Name = 'Fournisseurs - Achats'; Description = 'Évaluation des approvisionnements et services extérieurs'
            Questions = @(
                'Avons-nous une cartographie fournisseurs intégrant le risque RSE ?'
                'Suivez-vous la consommation par véhicule / par conducteur ?'
                'Quelle est l''origine de notre B100 (traçabilité, certifications durables) ?'
                'Avons-nous évalué la possibilité de souscrire à de l''énergie renouvelable ?'
                'Vos pneus usagés sont-ils recyclés ? Vos fournisseurs proposent-ils des alternatives recyclables ?'
                'Peut-on réduire le volume global d''emballages ?'
                'Connaissez-vous l''empreinte sociale et environnementale de vos sous-traitants ?'
                'Les assureurs valorisent-ils une flotte bas carbone ? Comparez-vous régulièrement les offres ?'
                'Les bâtiments loués respectent-ils les normes environnementales (DPE, HQE) ?'
                'Existe-t-il des alternatives plus responsables (coworking, logistique partagée) ?'
                'Existe-t-il une politique interne de mobilité durable ?'
            )
            Issues = @(
                'Concentration des achats sur quelques fournisseurs → dépendance'
                'Dépendance aux carburants fossiles (Gasoil, Essence, GAZ)'
                'Si B100, Difficulté à vérifier la traçabilité des huiles végétales utilisées'
                'Achats d''équipement à fort impact environnemental (Pneumatiques)'
                'Emballages parfois non recyclable et des quantités non suivies précisément'
                'Risque de dépendance d''un sous-traitant non écologique'
                'Sélection d''un assureur sans prise en compte des critères RSE'
                'Location de matériel ou locaux sans évaluation énergétique'
                'Dépenses élevées en déplacements (train, avion, voiture)'
            )
            Actions = @(
                'Déploiement d''une politique d''achats responsables (charte fournisseurs, critères ESG)'
                'Mettre en place un tableau de bord carburant (litres/km)'
                'Former les conducteurs à l''éco-conduite'
                'Diversifier l''approvisionnement en carburants alternatifs (BioGNV, électrique)'
                'Lancer un plan de renouvellement du parc léger'
                'Intégrer la rechapage/recyclage dans la politique d''achat'
                'Négocier la réduction des emballages superflus'
                'Cartographier les sous-traitants et intégrer des critères écologiques dans le choix'
                'Lancer un appel d''offres assurance intégrant des critères durables'
                'Négocier des bonus "flotte écoresponsable"'
                'Renégocier ou transférer vers des sites plus performants énergétiquement'
                'Réduire les locations temporaires par optimisation interne'
                'Mettre en place une politique de mobilité durable (priorité au train vs avion, encouragement au covoiturage, remboursement vélo)'
            )
        }
        @{
            Key = 'clients'; Name = 'Clients - Ventes'; Description = 'Analyse de la relation client et des ventes'
            Questions = @(
                'Quelle part du CA est réalisée avec des clients sensibles aux enjeux durables ?'
                'Vos clients demandent-ils des reporting carbone ?'
                'Pouvons-nous proposer une offre "durable" avec surcoût accepté ?'
                'Quelle part du CA est liée aux marchés publics (souvent sensibles au critère RSE) ?'
                'Ces remises sont-elles liées uniquement au volume, ou à des comportements vertueux (ex. engagement durable du client) ?'
                'Avons-nous recensé toutes les aides/subventions "durabilité" disponibles ?'
            )
            Issues = @(
                'Dépendance à quelques gros clients'
                'CA fortement dépendant d''une clientèle concentrée (risque de dépendance)'
                'Pas de différenciation tarifaire "bas carbone"'
                'Réductions commerciales peu reliées à une stratégie commerciale durable'
                'Subventions perçues (CEE, aides ADEME) non mises en avant'
            )
            Actions = @(
                'Développer une offre bas carbone pour fidéliser et conquérir les clients "verts"'
                'Développer une offre commerciale bas carbone'
                'Segmenter le portefeuille clients (clients sensibles au durable vs non sensibles)'
                'Conditionner certaines remises à des engagements RSE clients (utilisation de solutions bas carbone, fidélité dans les flux durables)'
                'Mettre en avant des bonus écologiques plutôt que des remises standard'
                'Suivi des subventions RSE perçues et à percevoir'
            )
        }
        @{
            Key = 'stocks'; Name = 'Stocks'; Description = 'Optimisation des stocks et circularité'
            Questions = @(
                'Constatez-vous un niveau de surstockage ou de pertes ?'
                'Avez-vous des solutions pour valoriser les invendus ?'
                'Travaillez-vous sur l''optimisation des volumes de stockage ?'
                'Avez-vous identifié des matières premières problématiques ?'
                'Quel est votre taux de rotation des stocks ?'
            )
            Issues = @(
                'Surstockage récurrent'
                'Matières premières à fort impact environnemental'
                'Gaspillage et invendus importants'
                'Obsolescence des produits stockés'
                'Espace de stockage mal optimisé'
            )
            Actions = @(
                'Diagnostic de l''empreinte écologique du stock'
                'Aide à la gestion prévisionnelle des approvisionnements'
                'Valorisation des invendus (réemploi, don, recyclage)'
                'Recommandation de produits alternatifs'
                'Optimisation des espaces de stockage'
            )
        }
        @{
            Key = 'deplacements'; Name = 'Déplacements - Mobilité'; Description = 'Optimisation des transports'
            Questions = @(
                'Avez-vous une politique interne de gestion des déplacements ?'
                'Quel % de votre flotte est électrique/hybride ?'
                'Les collaborateurs utilisent-ils des transports doux ?'
                'Avez-vous mesuré l''impact carbone de vos déplacements ?'
                'Le télétravail est-il pratiqué dans votre entreprise ?'
            )
            Issues = @(
                'Parc de véhicules diesel ancien'
                'Déplacements fréquents non optimisés'
                'Frais de déplacement très élevés'
                'Aucune politique de mobilité responsable'
                'Absence d''incitation aux transports doux'
            )
            Actions = @(
                'Évaluation de l''empreinte carbone des déplacements'
                'Plan de conversion véhicules électriques'
                'Mise en place d''un barème kilométrique vert'
                'Développement du télétravail'
                'Forfait mobilité durable pour les salariés'
            )
        }
        @{
            Key = 'energie_dechets'; Name = 'Énergie - Déchets - Eau'; Description = 'Gestion des ressources'
            Questions = @(
                'Disposez-vous de données sur vos consommations ?'
                'Avez-vous un suivi des déchets générés ?'
                'Quelles actions d''économie d''énergie avez-vous mis en place ?'
                'Quel est votre budget annuel énergie/eau ?'
                'Avez-vous des objectifs de réduction ?'
            )
            Issues = @(
                'Aucune mesure des consommations'
                'Coûts énergétiques croissants non maîtrisés'
                'Mauvaise gestion des déchets'
                'Absence de tri sélectif'
                'Consommation d''eau excessive'
            )
            Actions = @(
                'Mise en place d''un tableau de bord environnemental'
                'Partenariat avec un énergéticien spécialisé'
                'Optimisation de la facturation énergétique'
                'Accompagnement au tri et recyclage'
                'Installation de systèmes de mesure intelligents'
            )
        }
        @{
            Key = 'tresorerie_financement'; Name = 'Finance - Trésorerie'; Description = 'Leviers financiers verts'
            Questions = @(
                'Avez-vous bénéficié de financements verts ?'
                'Votre trésorerie est-elle placée de manière responsable ?'
                'Anticipez-vous des investissements liés à la transition ?'
                'Connaissez-vous les aides disponibles (ADEME, région...) ?'
                'Avez-vous évalué les risques climat sur votre modèle économique ?'
            )
            Issues = @(
                'Aucun recours aux financements durables'
                'Trésorerie inutilisée pour investissements verts'
                'Stratégie financière déconnectée de la durabilité'
                'Placements non alignés avec des critères ESG'
                'Manque de visibilité sur les aides disponibles'
            )
            Actions = @(
                'Identification des subventions vertes'
                'Recherche de financements à impact'
                'Simulation de ROI écologique'
                'Conseil en placements responsables'
                'Intégration des critères ESG dans la stratégie financière'
            )
        }
        @{
            Key = 'personnel'; Name = 'Personnel - RH'; Description = 'Politique RH durable'
            Questions = @(
                'Avez-vous intégré des critères RSE dans votre politique RH ?'
                'Proposez-vous des formations sur les enjeux environnementaux ?'
                'Quelles incitations aux comportements écoresponsables ?'
                'Le télétravail est-il pratiqué pour réduire les déplacements ?'
                'Avez-vous désigné un référent RSE ?'
            )
            Issues = @(
                'Absence de politique RH durable'
                'Salariés non formés aux enjeux environnementaux'
                'Mobilité domicile-travail non optimisée'
                'Aucune incitation aux bonnes pratiques'
                'Turn-over élevé dans les métiers verts'
            )
            Actions = @(
                'Élaboration d''une charte RH durable'
                'Programme de formation environnementale'
                'Mise en place d''un référent RSE'
                'Forfait mobilité durable'
                'Intégration des critères RSE dans les entretiens'
            )
        }
        @{
            Key = 'emballages'; Name = 'Emballages'; Description = 'Réduction des déchets d''emballage'
            Questions = @(
                'Quel pourcentage de vos emballages est recyclable ?'
                'Avez-vous étudié des solutions d''emballage réutilisables ?'
                'Subissez-vous des coûts liés à la taxe sur les emballages ?'
                'Vos clients exigent-ils des emballages écologiques ?'
                'Avez-vous calculé votre bilan carbone emballages ?'
            )
            Issues = @(
                'Emballages non recyclables'
                'Surcharges réglementaires'
                'Coûts de traitement élevés'
                'Image écologique dégradée'
                'Manque d''alternatives identifiées'
            )
            Actions = @(
                'Audit des flux d''emballages'
                'Optimisation des formats'
                'Recherche de solutions biosourcées'
                'Négociation avec les fournisseurs'
                'Calcul du ROI des alternatives'
            )
        }
    )
    GeneralQuestions = @(
        'Quelles sont vos principales préoccupations environnementales ?'
        'Avez-vous identifié des risques climatiques pour votre activité ?'
        'Disposez-vous d''une personne en charge des questions RSE ?'
        'Avez-vous établi un diagnostic de votre empreinte environnementale ?'
        'Quels seraient vos besoins prioritaires en accompagnement ?'
        'Avez-vous défini des objectifs de réduction d''impact ?'
    )
    CrossCuttingIssues = @(
        'Manque de données et indicateurs'
        'Absence de stratégie formalisée'
        'Réglementations mal anticipées'
        'Communication insuffisante en interne'
        'Budget dédié inexistant'
    )
    CrossCuttingActions = @(
        'Mise en place d''un système de mesure'
        'Élaboration d''une feuille de route RSE'
        'Veille réglementaire partagée'
        'Campagne de sensibilisation interne'
        'Création d''un budget transition'
    )
}
