@{
    Questions = @(
        @{
            Id = 1; Part = 1; Type = "radio"
            Text = "1 - Votre entreprise a-t-elle déjà engagé une démarche liée à l’environnement ou à la transition écologique ?"
            Description = ""
            Options = @(
                @{ Text = "Oui, dans le cadre d’une stratégie structurée"; Score = 4 }
                @{ Text = "Oui, ponctuellement mais sans formalisation"; Score = 3 }
                @{ Text = "Non, mais nous y réfléchissons"; Score = 2 }
                @{ Text = "Non, ce sujet n’est pas encore à l’ordre du jour"; Score = 0 }
            )
        }
        @{
            Id = 2; Part = 1; Type = "radio"
            Text = "2 - Avez-vous identifié des enjeux environnementaux propres à votre activité ?"
            Description = ""
            Options = @(
                @{ Text = "Oui, clairement identifiés"; Score = 4 }
                @{ Text = "Quelques éléments, mais non priorisés"; Score = 3 }
                @{ Text = "Pas vraiment"; Score = 1 }
                @{ Text = "Je ne sais pas"; Score = 0 }
            )
        }
        @{
            Id = 3; Part = 1; Type = "radio"
            Text = "3 - Votre entreprise communique-t-elle (même partiellement) sur ses engagements ou actions environnementales ?"
            Description = ""
            Options = @(
                @{ Text = "Oui, en interne et en externe"; Score = 4 }
                @{ Text = "Oui, uniquement en interne"; Score = 3 }
                @{ Text = "Non, mais nous y pensons"; Score = 2 }
                @{ Text = "Non, aucun besoin identifié"; Score = 0 }
            )
        }
        @{
            Id = 4; Part = 2; Type = "checkbox"
            Text = "4 - Quels sont, selon vous, les principaux freins à l'engagement de votre entreprise dans une démarche écologique ?"
            Description = "Plusieurs choix possibles"
            Options = @(
                @{ Text = "Le coût estimé"; Score = -1 }
                @{ Text = "Le manque de temps / de ressources humaines"; Score = -1 }
                @{ Text = "L’absence d’obligation réglementaire"; Score = -1 }
                @{ Text = "Le manque de compétences techniques ou d’accompagnement"; Score = -1 }
                @{ Text = "Je ne perçois pas l’intérêt immédiat"; Score = -1 }
            )
        }
        @{
            Id = 5; Part = 2; Type = "checkbox"
            Text = "5 - Quelles seraient vos motivations à engager une telle démarche ?"
            Description = "Plusieurs choix possibles"
            Options = @(
                @{ Text = "Respect des futures obligations réglementaires"; Score = 1 }
                @{ Text = "Réduction des coûts (énergie, déchets, etc.)"; Score = 1 }
                @{ Text = "Valorisation de l’image de l’entreprise"; Score = 1 }
                @{ Text = "Demandes de vos donneurs d’ordre / clients"; Score = 1 }
                @{ Text = "Accès à des financements ou des subventions"; Score = 1 }
            )
        }
        @{
            Id = 6; Part = 2; Type = "radio"
            Text = "6 - L’accompagnement de votre expert-comptable dans cette démarche vous semblerait-il utile ?"
            Description = ""
            Options = @(
                @{ Text = "Oui, je pense que c’est le bon interlocuteur"; Score = 3 }
                @{ Text = "Peut-être, s’il propose des outils adaptés"; Score = 2 }
                @{ Text = "Non, je préfère d’autres profils de conseil"; Score = 0 }
                @{ Text = "Je ne sais pas"; Score = 1 }
            )
        }
        @{
            Id = 7; Part = 3; Type = "rating"
            Text = "7 - Selon vous, votre expert-comptable est-il légitime pour vous accompagner sur les sujets suivants ?"
            Description = "Notez chaque item sur une échelle de 1 à 4 – 1 = Pas du tout légitime / 4 = Très légitime"
            Subjects = @(
                "Conseil en réduction des coûts énergétiques"
                "Intégration d’indicateurs environnementaux dans les tableaux de bord"
                "Analyse comptable pour identifier les leviers écologiques"
                "Aide au reporting extra-financier / durabilité"
                "Identification des aides publiques / subventions vertes"
            )
        }
        @{
            Id = 8; Part = 4; Type = "radio"
            Text = "8 - Seriez-vous prêt à bénéficier d’un accompagnement personnalisé du cabinet pour structurer une démarche environnementale ?"
            Description = ""
            Options = @(
                @{ Text = "Oui, dès maintenant"; Score = 4 }
                @{ Text = "Oui, à moyen terme"; Score = 3 }
                @{ Text = "Pas encore, mais j’y pense"; Score = 2 }
                @{ Text = "Non"; Score = 0 }
            )
        }
        @{
            Id = 9; Part = 4; Type = "radio"
            Text = "9 - Pensez-vous qu’il importe d’approfondir le sujet de votre transition écologique lors d’une prochaine discussion ?"
            Description = ""
            Options = @(
                @{ Text = "Oui"; Score = 1 }
                @{ Text = "Non"; Score = 0 }
            )
        }
    )
}
