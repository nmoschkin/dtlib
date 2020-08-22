

Namespace Info
    Public NotInheritable Class CountryCodeInfo

        Private _IsSatService As Boolean
        Private _Codes As New List(Of String)
        Private _PrimaryCodes As New List(Of String)
        Private _Name As String

        Private Shared _AllCountries As New List(Of CountryCodeInfo)
        Private Shared _AllSat As New List(Of CountryCodeInfo)

#Region "Shared country List Private Initialized Fields"

        Private Shared ReadOnly _Abkhazia As New CountryCodeInfo(False, "Abkhazia", "+7 840", "+7 940", "+995 44")
        Private Shared ReadOnly _Afghanistan As New CountryCodeInfo(False, "Afghanistan", "+93")
        Private Shared ReadOnly _ÅlandIslands As New CountryCodeInfo(False, "Åland Islands", "+358 18")
        Private Shared ReadOnly _Albania As New CountryCodeInfo(False, "Albania", "+355")
        Private Shared ReadOnly _Algeria As New CountryCodeInfo(False, "Algeria", "+213")
        Private Shared ReadOnly _AmericanSamoa As New CountryCodeInfo(False, "American Samoa", "+1 684")
        Private Shared ReadOnly _Andorra As New CountryCodeInfo(False, "Andorra", "+376")
        Private Shared ReadOnly _Angola As New CountryCodeInfo(False, "Angola", "+244")
        Private Shared ReadOnly _Anguilla As New CountryCodeInfo(False, "Anguilla", "+1 264")
        Private Shared ReadOnly _AntiguaandBarbuda As New CountryCodeInfo(False, "Antigua and Barbuda", "+1 268")
        Private Shared ReadOnly _Argentina As New CountryCodeInfo(False, "Argentina", "+54")
        Private Shared ReadOnly _Armenia As New CountryCodeInfo(False, "Armenia", "+374")
        Private Shared ReadOnly _Aruba As New CountryCodeInfo(False, "Aruba", "+297")
        Private Shared ReadOnly _Ascension As New CountryCodeInfo(False, "Ascension", "+247")
        Private Shared ReadOnly _Australia As New CountryCodeInfo(False, "Australia", "+61")
        Private Shared ReadOnly _AustralianExternalTerritories As New CountryCodeInfo(False, "Australian External Territories", "+672")
        Private Shared ReadOnly _Austria As New CountryCodeInfo(False, "Austria", "+43")
        Private Shared ReadOnly _Azerbaijan As New CountryCodeInfo(False, "Azerbaijan", "+994")
        Private Shared ReadOnly _Bahamas As New CountryCodeInfo(False, "Bahamas", "+1 242")
        Private Shared ReadOnly _Bahrain As New CountryCodeInfo(False, "Bahrain", "+973")
        Private Shared ReadOnly _Bangladesh As New CountryCodeInfo(False, "Bangladesh", "+880")
        Private Shared ReadOnly _Barbados As New CountryCodeInfo(False, "Barbados", "+1 246")
        Private Shared ReadOnly _Barbuda As New CountryCodeInfo(False, "Barbuda", "+1 268")
        Private Shared ReadOnly _Belarus As New CountryCodeInfo(False, "Belarus", "+375")
        Private Shared ReadOnly _Belgium As New CountryCodeInfo(False, "Belgium", "+32")
        Private Shared ReadOnly _Belize As New CountryCodeInfo(False, "Belize", "+501")
        Private Shared ReadOnly _Benin As New CountryCodeInfo(False, "Benin", "+229")
        Private Shared ReadOnly _Bermuda As New CountryCodeInfo(False, "Bermuda", "+1 441")
        Private Shared ReadOnly _Bhutan As New CountryCodeInfo(False, "Bhutan", "+975")
        Private Shared ReadOnly _Bolivia As New CountryCodeInfo(False, "Bolivia", "+591")
        Private Shared ReadOnly _Bonaire As New CountryCodeInfo(False, "Bonaire", "+599 7")
        Private Shared ReadOnly _BosniaandHerzegovina As New CountryCodeInfo(False, "Bosnia and Herzegovina", "+387")
        Private Shared ReadOnly _Botswana As New CountryCodeInfo(False, "Botswana", "+267")
        Private Shared ReadOnly _Brazil As New CountryCodeInfo(False, "Brazil", "+55")
        Private Shared ReadOnly _BritishIndianOceanTerritory As New CountryCodeInfo(False, "British Indian Ocean Territory", "+246")
        Private Shared ReadOnly _BritishVirginIslands As New CountryCodeInfo(False, "British Virgin Islands", "+1 284")
        Private Shared ReadOnly _BruneiDarussalam As New CountryCodeInfo(False, "Brunei Darussalam", "+673")
        Private Shared ReadOnly _Bulgaria As New CountryCodeInfo(False, "Bulgaria", "+359")
        Private Shared ReadOnly _BurkinaFaso As New CountryCodeInfo(False, "Burkina Faso", "+226")
        Private Shared ReadOnly _Burma As New CountryCodeInfo(False, "Burma", "+95")
        Private Shared ReadOnly _Burundi As New CountryCodeInfo(False, "Burundi", "+257")
        Private Shared ReadOnly _Cambodia As New CountryCodeInfo(False, "Cambodia", "+855")
        Private Shared ReadOnly _Cameroon As New CountryCodeInfo(False, "Cameroon", "+237")
        Private Shared ReadOnly _Canada As New CountryCodeInfo(False, "Canada", "+1")
        Private Shared ReadOnly _CapeVerde As New CountryCodeInfo(False, "Cape Verde", "+238")
        Private Shared ReadOnly _CaribbeanNetherlands As New CountryCodeInfo(False, "Caribbean Netherlands", "+599 3", "+599 4", "+599 7")
        Private Shared ReadOnly _CaymanIslands As New CountryCodeInfo(False, "Cayman Islands", "+1 345")
        Private Shared ReadOnly _CentralAfricanRepublic As New CountryCodeInfo(False, "Central African Republic", "+236")
        Private Shared ReadOnly _Chad As New CountryCodeInfo(False, "Chad", "+235")
        Private Shared ReadOnly _ChathamIsland As New CountryCodeInfo(False, "Chatham Island", "+64")
        Private Shared ReadOnly _Chile As New CountryCodeInfo(False, "Chile", "+56")
        Private Shared ReadOnly _China As New CountryCodeInfo(False, "China", "+86")
        Private Shared ReadOnly _ChristmasIsland As New CountryCodeInfo(False, "Christmas Island", "+61")
        Private Shared ReadOnly _Cocos As New CountryCodeInfo(False, "Cocos (Keeling) Islands", "+61")
        Private Shared ReadOnly _Colombia As New CountryCodeInfo(False, "Colombia", "+57")
        Private Shared ReadOnly _Comoros As New CountryCodeInfo(False, "Comoros", "+269")
        Private Shared ReadOnly _Congo As New CountryCodeInfo(False, "Congo", "+242")
        Private Shared ReadOnly _DemocraticRepublicoftheCongo As New CountryCodeInfo(False, "Democratic Republic of the Congo (Zaire)", "+243")
        Private Shared ReadOnly _CookIslands As New CountryCodeInfo(False, "Cook Islands", "+682")
        Private Shared ReadOnly _CostaRica As New CountryCodeInfo(False, "Costa Rica", "+506")
        Private Shared ReadOnly _CôtedIvoire As New CountryCodeInfo(False, "Côte d'Ivoire", "+225")
        Private Shared ReadOnly _Croatia As New CountryCodeInfo(False, "Croatia", "+385")
        Private Shared ReadOnly _Cuba As New CountryCodeInfo(False, "Cuba", "+53")
        Private Shared ReadOnly _GuantanamoBay As New CountryCodeInfo(False, "Guantanamo Bay", "+53 99")
        Private Shared ReadOnly _Curaçao As New CountryCodeInfo(False, "Curaçao", "+599 9")
        Private Shared ReadOnly _Cyprus As New CountryCodeInfo(False, "Cyprus", "+357")
        Private Shared ReadOnly _CzechRepublic As New CountryCodeInfo(False, "Czech Republic", "+420")
        Private Shared ReadOnly _Denmark As New CountryCodeInfo(False, "Denmark", "+45")
        Private Shared ReadOnly _DiegoGarcia As New CountryCodeInfo(False, "Diego Garcia", "+246")
        Private Shared ReadOnly _Djibouti As New CountryCodeInfo(False, "Djibouti", "+253")
        Private Shared ReadOnly _Dominica As New CountryCodeInfo(False, "Dominica", "+1 767")
        Private Shared ReadOnly _DominicanRepublic As New CountryCodeInfo(False, "Dominican Republic", "+1 809", "+1 829", "+1 849")
        Private Shared ReadOnly _EastTimor As New CountryCodeInfo(False, "East Timor", "+670")
        Private Shared ReadOnly _EasterIsland As New CountryCodeInfo(False, "Easter Island", "+56")
        Private Shared ReadOnly _Ecuador As New CountryCodeInfo(False, "Ecuador", "+593")
        Private Shared ReadOnly _Egypt As New CountryCodeInfo(False, "Egypt", "+20")
        Private Shared ReadOnly _ElSalvador As New CountryCodeInfo(False, "El Salvador", "+503")
        Private Shared ReadOnly _Ellipso As New CountryCodeInfo(True, "Ellipso (Mobile Satellite Service)", "+881 2", "+881 3")
        Private Shared ReadOnly _EMSAT As New CountryCodeInfo(True, "EMSAT (Mobile Satellite Service)", "+882 13")
        Private Shared ReadOnly _EquatorialGuinea As New CountryCodeInfo(False, "Equatorial Guinea", "+240")
        Private Shared ReadOnly _Eritrea As New CountryCodeInfo(False, "Eritrea", "+291")
        Private Shared ReadOnly _Estonia As New CountryCodeInfo(False, "Estonia", "+372")
        Private Shared ReadOnly _Ethiopia As New CountryCodeInfo(False, "Ethiopia", "+251")
        Private Shared ReadOnly _FalklandIslands As New CountryCodeInfo(False, "Falkland Islands", "+500")
        Private Shared ReadOnly _FaroeIslands As New CountryCodeInfo(False, "Faroe Islands", "+298")
        Private Shared ReadOnly _Fiji As New CountryCodeInfo(False, "Fiji", "+679")
        Private Shared ReadOnly _Finland As New CountryCodeInfo(False, "Finland", "+358")
        Private Shared ReadOnly _France As New CountryCodeInfo(False, "France", "+33")
        Private Shared ReadOnly _FrenchAntilles As New CountryCodeInfo(False, "French Antilles", "+596")
        Private Shared ReadOnly _FrenchGuiana As New CountryCodeInfo(False, "French Guiana", "+594")
        Private Shared ReadOnly _FrenchPolynesia As New CountryCodeInfo(False, "French Polynesia", "+689")
        Private Shared ReadOnly _Gabon As New CountryCodeInfo(False, "Gabon", "+241")
        Private Shared ReadOnly _Gambia As New CountryCodeInfo(False, "Gambia", "+220")
        Private Shared ReadOnly _Georgia As New CountryCodeInfo(False, "Georgia", "+995")
        Private Shared ReadOnly _Germany As New CountryCodeInfo(False, "Germany", "+49")
        Private Shared ReadOnly _Ghana As New CountryCodeInfo(False, "Ghana", "+233")
        Private Shared ReadOnly _Gibraltar As New CountryCodeInfo(False, "Gibraltar", "+350")
        Private Shared ReadOnly _GlobalMobileSatelliteSystem As New CountryCodeInfo(True, "Global Mobile Satellite System (GMSS)", "+881")
        Private Shared ReadOnly _Globalstar As New CountryCodeInfo(True, "Globalstar (Mobile Satellite Service)", "+881 8", "+881 9")
        Private Shared ReadOnly _Greece As New CountryCodeInfo(False, "Greece", "+30")
        Private Shared ReadOnly _Greenland As New CountryCodeInfo(False, "Greenland", "+299")
        Private Shared ReadOnly _Grenada As New CountryCodeInfo(False, "Grenada", "+1 473")
        Private Shared ReadOnly _Guadeloupe As New CountryCodeInfo(False, "Guadeloupe", "+590")
        Private Shared ReadOnly _Guam As New CountryCodeInfo(False, "Guam", "+1 671")
        Private Shared ReadOnly _Guatemala As New CountryCodeInfo(False, "Guatemala", "+502")
        Private Shared ReadOnly _Guernsey As New CountryCodeInfo(False, "Guernsey", "+44")
        Private Shared ReadOnly _Guinea As New CountryCodeInfo(False, "Guinea", "+224")
        Private Shared ReadOnly _GuineaBissau As New CountryCodeInfo(False, "Guinea-Bissau", "+245")
        Private Shared ReadOnly _Guyana As New CountryCodeInfo(False, "Guyana", "+592")
        Private Shared ReadOnly _Haiti As New CountryCodeInfo(False, "Haiti", "+509")
        Private Shared ReadOnly _Honduras As New CountryCodeInfo(False, "Honduras", "+504")
        Private Shared ReadOnly _HongKong As New CountryCodeInfo(False, "Hong Kong", "+852")
        Private Shared ReadOnly _Hungary As New CountryCodeInfo(False, "Hungary", "+36")
        Private Shared ReadOnly _Iceland As New CountryCodeInfo(False, "Iceland", "+354")
        Private Shared ReadOnly _ICOGlobal As New CountryCodeInfo(True, "ICO Global (Mobile Satellite Service)", "+881 0", "+881 1")
        Private Shared ReadOnly _India As New CountryCodeInfo(False, "India", "+91")
        Private Shared ReadOnly _Indonesia As New CountryCodeInfo(False, "Indonesia", "+62")
        Private Shared ReadOnly _InmarsatSNAC As New CountryCodeInfo(False, "Inmarsat SNAC", "+870")
        Private Shared ReadOnly _InternationalFreephoneService As New CountryCodeInfo(True, "International Freephone Service", "+800")
        Private Shared ReadOnly _InternationalSharedCostService As New CountryCodeInfo(True, "International Shared ReadOnly  Cost Service (ISCS)", "+808")
        Private Shared ReadOnly _Iran As New CountryCodeInfo(False, "Iran", "+98")
        Private Shared ReadOnly _Iraq As New CountryCodeInfo(False, "Iraq", "+964")
        Private Shared ReadOnly _Ireland As New CountryCodeInfo(False, "Ireland", "+353")
        Private Shared ReadOnly _Iridium As New CountryCodeInfo(True, "Iridium (Mobile Satellite Service)", "+881 6", "+881 7")
        Private Shared ReadOnly _IsleofMan As New CountryCodeInfo(False, "Isle of Man", "+44")
        Private Shared ReadOnly _Israel As New CountryCodeInfo(False, "Israel", "+972")
        Private Shared ReadOnly _Italy As New CountryCodeInfo(False, "Italy", "+39")
        Private Shared ReadOnly _Jamaica As New CountryCodeInfo(False, "Jamaica", "+1 876")
        Private Shared ReadOnly _JanMayen As New CountryCodeInfo(False, "Jan Mayen", "+47 79")
        Private Shared ReadOnly _Japan As New CountryCodeInfo(False, "Japan", "+81")
        Private Shared ReadOnly _Jersey As New CountryCodeInfo(False, "Jersey", "+44")
        Private Shared ReadOnly _Jordan As New CountryCodeInfo(False, "Jordan", "+962")
        Private Shared ReadOnly _Kazakhstan As New CountryCodeInfo(False, "Kazakhstan", "+7 6", "+7 7")
        Private Shared ReadOnly _Kenya As New CountryCodeInfo(False, "Kenya", "+254")
        Private Shared ReadOnly _Kiribati As New CountryCodeInfo(False, "Kiribati", "+686")
        Private Shared ReadOnly _NorthKorea As New CountryCodeInfo(False, "North Korea", "+850")
        Private Shared ReadOnly _SouthKorea As New CountryCodeInfo(False, "South Korea", "+82")
        Private Shared ReadOnly _Kuwait As New CountryCodeInfo(False, "Kuwait", "+965")
        Private Shared ReadOnly _Kyrgyzstan As New CountryCodeInfo(False, "Kyrgyzstan", "+996")
        Private Shared ReadOnly _Laos As New CountryCodeInfo(False, "Laos", "+856")
        Private Shared ReadOnly _Latvia As New CountryCodeInfo(False, "Latvia", "+371")
        Private Shared ReadOnly _Lebanon As New CountryCodeInfo(False, "Lebanon", "+961")
        Private Shared ReadOnly _Lesotho As New CountryCodeInfo(False, "Lesotho", "+266")
        Private Shared ReadOnly _Liberia As New CountryCodeInfo(False, "Liberia", "+231")
        Private Shared ReadOnly _Libya As New CountryCodeInfo(False, "Libya", "+218")
        Private Shared ReadOnly _Liechtenstein As New CountryCodeInfo(False, "Liechtenstein", "+423")
        Private Shared ReadOnly _Lithuania As New CountryCodeInfo(False, "Lithuania", "+370")
        Private Shared ReadOnly _Luxembourg As New CountryCodeInfo(False, "Luxembourg", "+352")
        Private Shared ReadOnly _Macau As New CountryCodeInfo(False, "Macau", "+853")
        Private Shared ReadOnly _Macedonia As New CountryCodeInfo(False, "Macedonia", "+389")
        Private Shared ReadOnly _Madagascar As New CountryCodeInfo(False, "Madagascar", "+261")
        Private Shared ReadOnly _Malawi As New CountryCodeInfo(False, "Malawi", "+265")
        Private Shared ReadOnly _Malaysia As New CountryCodeInfo(False, "Malaysia", "+60")
        Private Shared ReadOnly _Maldives As New CountryCodeInfo(False, "Maldives", "+960")
        Private Shared ReadOnly _Mali As New CountryCodeInfo(False, "Mali", "+223")
        Private Shared ReadOnly _Malta As New CountryCodeInfo(False, "Malta", "+356")
        Private Shared ReadOnly _MarshallIslands As New CountryCodeInfo(False, "Marshall Islands", "+692")
        Private Shared ReadOnly _Martinique As New CountryCodeInfo(False, "Martinique", "+596")
        Private Shared ReadOnly _Mauritania As New CountryCodeInfo(False, "Mauritania", "+222")
        Private Shared ReadOnly _Mauritius As New CountryCodeInfo(False, "Mauritius", "+230")
        Private Shared ReadOnly _Mayotte As New CountryCodeInfo(False, "Mayotte", "+262")
        Private Shared ReadOnly _Mexico As New CountryCodeInfo(False, "Mexico", "+52")
        Private Shared ReadOnly _FederatedStatesofMicronesia As New CountryCodeInfo(False, "Federated States of Micronesia", "+691")
        Private Shared ReadOnly _MidwayIsland As New CountryCodeInfo(False, "Midway Island USA", "+1 808")
        Private Shared ReadOnly _Moldova As New CountryCodeInfo(False, "Moldova", "+373")
        Private Shared ReadOnly _Monaco As New CountryCodeInfo(False, "Monaco", "+377")
        Private Shared ReadOnly _Mongolia As New CountryCodeInfo(False, "Mongolia", "+976")
        Private Shared ReadOnly _Montenegro As New CountryCodeInfo(False, "Montenegro", "+382")
        Private Shared ReadOnly _Montserrat As New CountryCodeInfo(False, "Montserrat", "+1 664")
        Private Shared ReadOnly _Morocco As New CountryCodeInfo(False, "Morocco", "+212")
        Private Shared ReadOnly _Mozambique As New CountryCodeInfo(False, "Mozambique", "+258")
        Private Shared ReadOnly _Namibia As New CountryCodeInfo(False, "Namibia", "+264")
        Private Shared ReadOnly _Nauru As New CountryCodeInfo(False, "Nauru", "+674")
        Private Shared ReadOnly _Nepal As New CountryCodeInfo(False, "Nepal", "+977")
        Private Shared ReadOnly _Netherlands As New CountryCodeInfo(False, "Netherlands", "+31")
        Private Shared ReadOnly _Nevis As New CountryCodeInfo(False, "Nevis", "+1 869")
        Private Shared ReadOnly _NewCaledonia As New CountryCodeInfo(False, "New Caledonia", "+687")
        Private Shared ReadOnly _NewZealand As New CountryCodeInfo(False, "New Zealand", "+64")
        Private Shared ReadOnly _Nicaragua As New CountryCodeInfo(False, "Nicaragua", "+505")
        Private Shared ReadOnly _Niger As New CountryCodeInfo(False, "Niger", "+227")
        Private Shared ReadOnly _Nigeria As New CountryCodeInfo(False, "Nigeria", "+234")
        Private Shared ReadOnly _Niue As New CountryCodeInfo(False, "Niue", "+683")
        Private Shared ReadOnly _NorfolkIsland As New CountryCodeInfo(False, "Norfolk Island", "+672")
        Private Shared ReadOnly _NorthernMarianaIslands As New CountryCodeInfo(False, "Northern Mariana Islands", "+1 670")
        Private Shared ReadOnly _Norway As New CountryCodeInfo(False, "Norway", "+47")
        Private Shared ReadOnly _Oman As New CountryCodeInfo(False, "Oman", "+968")
        Private Shared ReadOnly _Pakistan As New CountryCodeInfo(False, "Pakistan", "+92")
        Private Shared ReadOnly _Palau As New CountryCodeInfo(False, "Palau", "+680")
        Private Shared ReadOnly _PalestinianTerritories As New CountryCodeInfo(False, "Palestinian Territories", "+970")
        Private Shared ReadOnly _Panama As New CountryCodeInfo(False, "Panama", "+507")
        Private Shared ReadOnly _PapuaNewGuinea As New CountryCodeInfo(False, "Papua New Guinea", "+675")
        Private Shared ReadOnly _Paraguay As New CountryCodeInfo(False, "Paraguay", "+595")
        Private Shared ReadOnly _Peru As New CountryCodeInfo(False, "Peru", "+51")
        Private Shared ReadOnly _Philippines As New CountryCodeInfo(False, "Philippines", "+63")
        Private Shared ReadOnly _PitcairnIslands As New CountryCodeInfo(False, "Pitcairn Islands", "+64")
        Private Shared ReadOnly _Poland As New CountryCodeInfo(False, "Poland", "+48")
        Private Shared ReadOnly _Portugal As New CountryCodeInfo(False, "Portugal", "+351")
        Private Shared ReadOnly _PuertoRico As New CountryCodeInfo(False, "Puerto Rico", "+1 787", "+1 939")
        Private Shared ReadOnly _Qatar As New CountryCodeInfo(False, "Qatar", "+974")
        Private Shared ReadOnly _Réunion As New CountryCodeInfo(False, "Réunion", "+262")
        Private Shared ReadOnly _Romania As New CountryCodeInfo(False, "Romania", "+40")
        Private Shared ReadOnly _Russia As New CountryCodeInfo(False, "Russia", "+7")
        Private Shared ReadOnly _Rwanda As New CountryCodeInfo(False, "Rwanda", "+250")
        Private Shared ReadOnly _Saba As New CountryCodeInfo(False, "Saba", "+599 4")
        Private Shared ReadOnly _SaintBarthélemy As New CountryCodeInfo(False, "Saint Barthélemy", "+590")
        Private Shared ReadOnly _SaintHelena As New CountryCodeInfo(False, "Saint Helena", "+290")
        Private Shared ReadOnly _SaintKittsandNevis As New CountryCodeInfo(False, "Saint Kitts and Nevis", "+1 869")
        Private Shared ReadOnly _SaintLucia As New CountryCodeInfo(False, "Saint Lucia", "+1 758")
        Private Shared ReadOnly _SaintMartin As New CountryCodeInfo(False, "Saint Martin (France)", "+590")
        Private Shared ReadOnly _SaintPierreandMiquelon As New CountryCodeInfo(False, "Saint Pierre and Miquelon", "+508")
        Private Shared ReadOnly _SaintVincentandtheGrenadines As New CountryCodeInfo(False, "Saint Vincent and the Grenadines", "+1 784")
        Private Shared ReadOnly _Samoa As New CountryCodeInfo(False, "Samoa", "+685")
        Private Shared ReadOnly _SanMarino As New CountryCodeInfo(False, "San Marino", "+378")
        Private Shared ReadOnly _SãoToméandPríncipe As New CountryCodeInfo(False, "São Tomé and Príncipe", "+239")
        Private Shared ReadOnly _SaudiArabia As New CountryCodeInfo(False, "Saudi Arabia", "+966")
        Private Shared ReadOnly _Senegal As New CountryCodeInfo(False, "Senegal", "+221")
        Private Shared ReadOnly _Serbia As New CountryCodeInfo(False, "Serbia", "+381")
        Private Shared ReadOnly _Seychelles As New CountryCodeInfo(False, "Seychelles", "+248")
        Private Shared ReadOnly _SierraLeone As New CountryCodeInfo(False, "Sierra Leone", "+232")
        Private Shared ReadOnly _Singapore As New CountryCodeInfo(False, "Singapore", "+65")
        Private Shared ReadOnly _SintEustatius As New CountryCodeInfo(False, "Sint Eustatius", "+599 3")
        Private Shared ReadOnly _SintMaarten As New CountryCodeInfo(False, "Sint Maarten (Netherlands)", "+1 721")
        Private Shared ReadOnly _Slovakia As New CountryCodeInfo(False, "Slovakia", "+421")
        Private Shared ReadOnly _Slovenia As New CountryCodeInfo(False, "Slovenia", "+386")
        Private Shared ReadOnly _SolomonIslands As New CountryCodeInfo(False, "Solomon Islands", "+677")
        Private Shared ReadOnly _Somalia As New CountryCodeInfo(False, "Somalia", "+252")
        Private Shared ReadOnly _SouthAfrica As New CountryCodeInfo(False, "South Africa", "+27")
        Private Shared ReadOnly _SouthGeorgiaandtheSouthSandwichIslands As New CountryCodeInfo(False, "South Georgia and the South Sandwich Islands", "+500")
        Private Shared ReadOnly _SouthOssetia As New CountryCodeInfo(False, "South Ossetia", "+995 34")
        Private Shared ReadOnly _SouthSudan As New CountryCodeInfo(False, "South Sudan", "+211")
        Private Shared ReadOnly _Spain As New CountryCodeInfo(False, "Spain", "+34")
        Private Shared ReadOnly _SriLanka As New CountryCodeInfo(False, "Sri Lanka", "+94")
        Private Shared ReadOnly _Sudan As New CountryCodeInfo(False, "Sudan", "+249")
        Private Shared ReadOnly _Suriname As New CountryCodeInfo(False, "Suriname", "+597")
        Private Shared ReadOnly _Svalbard As New CountryCodeInfo(False, "Svalbard", "+47 79")
        Private Shared ReadOnly _Swaziland As New CountryCodeInfo(False, "Swaziland", "+268")
        Private Shared ReadOnly _Sweden As New CountryCodeInfo(False, "Sweden", "+46")
        Private Shared ReadOnly _Switzerland As New CountryCodeInfo(False, "Switzerland", "+41")
        Private Shared ReadOnly _Syria As New CountryCodeInfo(False, "Syria", "+963")
        Private Shared ReadOnly _Taiwan As New CountryCodeInfo(False, "Taiwan", "+886")
        Private Shared ReadOnly _Tajikistan As New CountryCodeInfo(False, "Tajikistan", "+992")
        Private Shared ReadOnly _Tanzania As New CountryCodeInfo(False, "Tanzania", "+255")
        Private Shared ReadOnly _Thailand As New CountryCodeInfo(False, "Thailand", "+66")
        Private Shared ReadOnly _Thuraya As New CountryCodeInfo(True, "Thuraya (Mobile Satellite Service)", "+882 16")
        Private Shared ReadOnly _Togo As New CountryCodeInfo(False, "Togo", "+228")
        Private Shared ReadOnly _Tokelau As New CountryCodeInfo(False, "Tokelau", "+690")
        Private Shared ReadOnly _Tonga As New CountryCodeInfo(False, "Tonga", "+676")
        Private Shared ReadOnly _TrinidadandTobago As New CountryCodeInfo(False, "Trinidad and Tobago", "+1 868")
        Private Shared ReadOnly _TristandaCunha As New CountryCodeInfo(False, "Tristan da Cunha", "+290 8")
        Private Shared ReadOnly _Tunisia As New CountryCodeInfo(False, "Tunisia", "+216")
        Private Shared ReadOnly _Turkey As New CountryCodeInfo(False, "Turkey", "+90")
        Private Shared ReadOnly _Turkmenistan As New CountryCodeInfo(False, "Turkmenistan", "+993")
        Private Shared ReadOnly _TurksandCaicosIslands As New CountryCodeInfo(False, "Turks and Caicos Islands", "+1 649")
        Private Shared ReadOnly _Tuvalu As New CountryCodeInfo(False, "Tuvalu", "+688")
        Private Shared ReadOnly _Uganda As New CountryCodeInfo(False, "Uganda", "+256")
        Private Shared ReadOnly _Ukraine As New CountryCodeInfo(False, "Ukraine", "+380")
        Private Shared ReadOnly _UnitedArabEmirates As New CountryCodeInfo(False, "United Arab Emirates", "+971")
        Private Shared ReadOnly _UnitedKingdom As New CountryCodeInfo(False, "United Kingdom", "+44")
        Private Shared ReadOnly _UnitedStates As New CountryCodeInfo(False, "United States", "+1")
        Private Shared ReadOnly _UniversalPersonalTelecommunications As New CountryCodeInfo(True, "Universal Personal Telecommunications (UPT)", "+878")
        Private Shared ReadOnly _Uruguay As New CountryCodeInfo(False, "Uruguay", "+598")
        Private Shared ReadOnly _USVirginIslands As New CountryCodeInfo(False, "US Virgin Islands", "+1 340")
        Private Shared ReadOnly _Uzbekistan As New CountryCodeInfo(False, "Uzbekistan", "+998")
        Private Shared ReadOnly _Vanuatu As New CountryCodeInfo(False, "Vanuatu", "+678")
        Private Shared ReadOnly _Venezuela As New CountryCodeInfo(False, "Venezuela", "+58")
        Private Shared ReadOnly _VaticanCityState As New CountryCodeInfo(False, "Vatican City State (Holy See)", "+39 06 698", "+379")
        Private Shared ReadOnly _Vietnam As New CountryCodeInfo(False, "Vietnam", "+84")
        Private Shared ReadOnly _WakeIsland As New CountryCodeInfo(False, "Wake Island USA", "+1 808")
        Private Shared ReadOnly _WallisandFutuna As New CountryCodeInfo(False, "Wallis and Futuna", "+681")
        Private Shared ReadOnly _Yemen As New CountryCodeInfo(False, "Yemen", "+967")
        Private Shared ReadOnly _Zambia As New CountryCodeInfo(False, "Zambia", "+260")
        Private Shared ReadOnly _Zanzibar As New CountryCodeInfo(False, "Zanzibar", "+255")
        Private Shared ReadOnly _Zimbabwe As New CountryCodeInfo(False, "Zimbabwe", "+268")

#End Region

#Region "Shared Public Country Properties"

        Public Shared ReadOnly Property Abkhazia As CountryCodeInfo
            Get
                Return _Abkhazia
            End Get
        End Property

        Public Shared ReadOnly Property Afghanistan As CountryCodeInfo
            Get
                Return _Afghanistan
            End Get
        End Property

        Public Shared ReadOnly Property ÅlandIslands As CountryCodeInfo
            Get
                Return _ÅlandIslands
            End Get
        End Property

        Public Shared ReadOnly Property Albania As CountryCodeInfo
            Get
                Return _Albania
            End Get
        End Property

        Public Shared ReadOnly Property Algeria As CountryCodeInfo
            Get
                Return _Algeria
            End Get
        End Property

        Public Shared ReadOnly Property AmericanSamoa As CountryCodeInfo
            Get
                Return _AmericanSamoa
            End Get
        End Property

        Public Shared ReadOnly Property Andorra As CountryCodeInfo
            Get
                Return _Andorra
            End Get
        End Property

        Public Shared ReadOnly Property Angola As CountryCodeInfo
            Get
                Return _Angola
            End Get
        End Property

        Public Shared ReadOnly Property Anguilla As CountryCodeInfo
            Get
                Return _Anguilla
            End Get
        End Property

        Public Shared ReadOnly Property AntiguaandBarbuda As CountryCodeInfo
            Get
                Return _AntiguaandBarbuda
            End Get
        End Property

        Public Shared ReadOnly Property Argentina As CountryCodeInfo
            Get
                Return _Argentina
            End Get
        End Property

        Public Shared ReadOnly Property Armenia As CountryCodeInfo
            Get
                Return _Armenia
            End Get
        End Property

        Public Shared ReadOnly Property Aruba As CountryCodeInfo
            Get
                Return _Aruba
            End Get
        End Property

        Public Shared ReadOnly Property Ascension As CountryCodeInfo
            Get
                Return _Ascension
            End Get
        End Property

        Public Shared ReadOnly Property Australia As CountryCodeInfo
            Get
                Return _Australia
            End Get
        End Property

        Public Shared ReadOnly Property AustralianExternalTerritories As CountryCodeInfo
            Get
                Return _AustralianExternalTerritories
            End Get
        End Property

        Public Shared ReadOnly Property Austria As CountryCodeInfo
            Get
                Return _Austria
            End Get
        End Property

        Public Shared ReadOnly Property Azerbaijan As CountryCodeInfo
            Get
                Return _Azerbaijan
            End Get
        End Property

        Public Shared ReadOnly Property Bahamas As CountryCodeInfo
            Get
                Return _Bahamas
            End Get
        End Property

        Public Shared ReadOnly Property Bahrain As CountryCodeInfo
            Get
                Return _Bahrain
            End Get
        End Property

        Public Shared ReadOnly Property Bangladesh As CountryCodeInfo
            Get
                Return _Bangladesh
            End Get
        End Property

        Public Shared ReadOnly Property Barbados As CountryCodeInfo
            Get
                Return _Barbados
            End Get
        End Property

        Public Shared ReadOnly Property Barbuda As CountryCodeInfo
            Get
                Return _Barbuda
            End Get
        End Property

        Public Shared ReadOnly Property Belarus As CountryCodeInfo
            Get
                Return _Belarus
            End Get
        End Property

        Public Shared ReadOnly Property Belgium As CountryCodeInfo
            Get
                Return _Belgium
            End Get
        End Property

        Public Shared ReadOnly Property Belize As CountryCodeInfo
            Get
                Return _Belize
            End Get
        End Property

        Public Shared ReadOnly Property Benin As CountryCodeInfo
            Get
                Return _Benin
            End Get
        End Property

        Public Shared ReadOnly Property Bermuda As CountryCodeInfo
            Get
                Return _Bermuda
            End Get
        End Property

        Public Shared ReadOnly Property Bhutan As CountryCodeInfo
            Get
                Return _Bhutan
            End Get
        End Property

        Public Shared ReadOnly Property Bolivia As CountryCodeInfo
            Get
                Return _Bolivia
            End Get
        End Property

        Public Shared ReadOnly Property Bonaire As CountryCodeInfo
            Get
                Return _Bonaire
            End Get
        End Property

        Public Shared ReadOnly Property BosniaandHerzegovina As CountryCodeInfo
            Get
                Return _BosniaandHerzegovina
            End Get
        End Property

        Public Shared ReadOnly Property Botswana As CountryCodeInfo
            Get
                Return _Botswana
            End Get
        End Property

        Public Shared ReadOnly Property Brazil As CountryCodeInfo
            Get
                Return _Brazil
            End Get
        End Property

        Public Shared ReadOnly Property BritishIndianOceanTerritory As CountryCodeInfo
            Get
                Return _BritishIndianOceanTerritory
            End Get
        End Property

        Public Shared ReadOnly Property BritishVirginIslands As CountryCodeInfo
            Get
                Return _BritishVirginIslands
            End Get
        End Property

        Public Shared ReadOnly Property BruneiDarussalam As CountryCodeInfo
            Get
                Return _BruneiDarussalam
            End Get
        End Property

        Public Shared ReadOnly Property Bulgaria As CountryCodeInfo
            Get
                Return _Bulgaria
            End Get
        End Property

        Public Shared ReadOnly Property BurkinaFaso As CountryCodeInfo
            Get
                Return _BurkinaFaso
            End Get
        End Property

        Public Shared ReadOnly Property Burma As CountryCodeInfo
            Get
                Return _Burma
            End Get
        End Property

        Public Shared ReadOnly Property Burundi As CountryCodeInfo
            Get
                Return _Burundi
            End Get
        End Property

        Public Shared ReadOnly Property Cambodia As CountryCodeInfo
            Get
                Return _Cambodia
            End Get
        End Property

        Public Shared ReadOnly Property Cameroon As CountryCodeInfo
            Get
                Return _Cameroon
            End Get
        End Property

        Public Shared ReadOnly Property Canada As CountryCodeInfo
            Get
                Return _Canada
            End Get
        End Property

        Public Shared ReadOnly Property CapeVerde As CountryCodeInfo
            Get
                Return _CapeVerde
            End Get
        End Property

        Public Shared ReadOnly Property CaribbeanNetherlands As CountryCodeInfo
            Get
                Return _CaribbeanNetherlands
            End Get
        End Property

        Public Shared ReadOnly Property CaymanIslands As CountryCodeInfo
            Get
                Return _CaymanIslands
            End Get
        End Property

        Public Shared ReadOnly Property CentralAfricanRepublic As CountryCodeInfo
            Get
                Return _CentralAfricanRepublic
            End Get
        End Property

        Public Shared ReadOnly Property Chad As CountryCodeInfo
            Get
                Return _Chad
            End Get
        End Property

        Public Shared ReadOnly Property ChathamIsland As CountryCodeInfo
            Get
                Return _ChathamIsland
            End Get
        End Property

        Public Shared ReadOnly Property Chile As CountryCodeInfo
            Get
                Return _Chile
            End Get
        End Property

        Public Shared ReadOnly Property China As CountryCodeInfo
            Get
                Return _China
            End Get
        End Property

        Public Shared ReadOnly Property ChristmasIsland As CountryCodeInfo
            Get
                Return _ChristmasIsland
            End Get
        End Property

        Public Shared ReadOnly Property Cocos As CountryCodeInfo
            Get
                Return _Cocos
            End Get
        End Property

        Public Shared ReadOnly Property Colombia As CountryCodeInfo
            Get
                Return _Colombia
            End Get
        End Property

        Public Shared ReadOnly Property Comoros As CountryCodeInfo
            Get
                Return _Comoros
            End Get
        End Property

        Public Shared ReadOnly Property Congo As CountryCodeInfo
            Get
                Return _Congo
            End Get
        End Property

        Public Shared ReadOnly Property DemocraticRepublicoftheCongo As CountryCodeInfo
            Get
                Return _DemocraticRepublicoftheCongo
            End Get
        End Property

        Public Shared ReadOnly Property CookIslands As CountryCodeInfo
            Get
                Return _CookIslands
            End Get
        End Property

        Public Shared ReadOnly Property CostaRica As CountryCodeInfo
            Get
                Return _CostaRica
            End Get
        End Property

        Public Shared ReadOnly Property CôtedIvoire As CountryCodeInfo
            Get
                Return _CôtedIvoire
            End Get
        End Property

        Public Shared ReadOnly Property Croatia As CountryCodeInfo
            Get
                Return _Croatia
            End Get
        End Property

        Public Shared ReadOnly Property Cuba As CountryCodeInfo
            Get
                Return _Cuba
            End Get
        End Property

        Public Shared ReadOnly Property GuantanamoBay As CountryCodeInfo
            Get
                Return _GuantanamoBay
            End Get
        End Property

        Public Shared ReadOnly Property Curaçao As CountryCodeInfo
            Get
                Return _Curaçao
            End Get
        End Property

        Public Shared ReadOnly Property Cyprus As CountryCodeInfo
            Get
                Return _Cyprus
            End Get
        End Property

        Public Shared ReadOnly Property CzechRepublic As CountryCodeInfo
            Get
                Return _CzechRepublic
            End Get
        End Property

        Public Shared ReadOnly Property Denmark As CountryCodeInfo
            Get
                Return _Denmark
            End Get
        End Property

        Public Shared ReadOnly Property DiegoGarcia As CountryCodeInfo
            Get
                Return _DiegoGarcia
            End Get
        End Property

        Public Shared ReadOnly Property Djibouti As CountryCodeInfo
            Get
                Return _Djibouti
            End Get
        End Property

        Public Shared ReadOnly Property Dominica As CountryCodeInfo
            Get
                Return _Dominica
            End Get
        End Property

        Public Shared ReadOnly Property DominicanRepublic As CountryCodeInfo
            Get
                Return _DominicanRepublic
            End Get
        End Property

        Public Shared ReadOnly Property EastTimor As CountryCodeInfo
            Get
                Return _EastTimor
            End Get
        End Property

        Public Shared ReadOnly Property EasterIsland As CountryCodeInfo
            Get
                Return _EasterIsland
            End Get
        End Property

        Public Shared ReadOnly Property Ecuador As CountryCodeInfo
            Get
                Return _Ecuador
            End Get
        End Property

        Public Shared ReadOnly Property Egypt As CountryCodeInfo
            Get
                Return _Egypt
            End Get
        End Property

        Public Shared ReadOnly Property ElSalvador As CountryCodeInfo
            Get
                Return _ElSalvador
            End Get
        End Property

        Public Shared ReadOnly Property Ellipso As CountryCodeInfo
            Get
                Return _Ellipso
            End Get
        End Property

        Public Shared ReadOnly Property EMSAT As CountryCodeInfo
            Get
                Return _EMSAT
            End Get
        End Property

        Public Shared ReadOnly Property EquatorialGuinea As CountryCodeInfo
            Get
                Return _EquatorialGuinea
            End Get
        End Property

        Public Shared ReadOnly Property Eritrea As CountryCodeInfo
            Get
                Return _Eritrea
            End Get
        End Property

        Public Shared ReadOnly Property Estonia As CountryCodeInfo
            Get
                Return _Estonia
            End Get
        End Property

        Public Shared ReadOnly Property Ethiopia As CountryCodeInfo
            Get
                Return _Ethiopia
            End Get
        End Property

        Public Shared ReadOnly Property FalklandIslands As CountryCodeInfo
            Get
                Return _FalklandIslands
            End Get
        End Property

        Public Shared ReadOnly Property FaroeIslands As CountryCodeInfo
            Get
                Return _FaroeIslands
            End Get
        End Property

        Public Shared ReadOnly Property Fiji As CountryCodeInfo
            Get
                Return _Fiji
            End Get
        End Property

        Public Shared ReadOnly Property Finland As CountryCodeInfo
            Get
                Return _Finland
            End Get
        End Property

        Public Shared ReadOnly Property France As CountryCodeInfo
            Get
                Return _France
            End Get
        End Property

        Public Shared ReadOnly Property FrenchAntilles As CountryCodeInfo
            Get
                Return _FrenchAntilles
            End Get
        End Property

        Public Shared ReadOnly Property FrenchGuiana As CountryCodeInfo
            Get
                Return _FrenchGuiana
            End Get
        End Property

        Public Shared ReadOnly Property FrenchPolynesia As CountryCodeInfo
            Get
                Return _FrenchPolynesia
            End Get
        End Property

        Public Shared ReadOnly Property Gabon As CountryCodeInfo
            Get
                Return _Gabon
            End Get
        End Property

        Public Shared ReadOnly Property Gambia As CountryCodeInfo
            Get
                Return _Gambia
            End Get
        End Property

        Public Shared ReadOnly Property Georgia As CountryCodeInfo
            Get
                Return _Georgia
            End Get
        End Property

        Public Shared ReadOnly Property Germany As CountryCodeInfo
            Get
                Return _Germany
            End Get
        End Property

        Public Shared ReadOnly Property Ghana As CountryCodeInfo
            Get
                Return _Ghana
            End Get
        End Property

        Public Shared ReadOnly Property Gibraltar As CountryCodeInfo
            Get
                Return _Gibraltar
            End Get
        End Property

        Public Shared ReadOnly Property GlobalMobileSatelliteSystem As CountryCodeInfo
            Get
                Return _GlobalMobileSatelliteSystem
            End Get
        End Property

        Public Shared ReadOnly Property Globalstar As CountryCodeInfo
            Get
                Return _Globalstar
            End Get
        End Property

        Public Shared ReadOnly Property Greece As CountryCodeInfo
            Get
                Return _Greece
            End Get
        End Property

        Public Shared ReadOnly Property Greenland As CountryCodeInfo
            Get
                Return _Greenland
            End Get
        End Property

        Public Shared ReadOnly Property Grenada As CountryCodeInfo
            Get
                Return _Grenada
            End Get
        End Property

        Public Shared ReadOnly Property Guadeloupe As CountryCodeInfo
            Get
                Return _Guadeloupe
            End Get
        End Property

        Public Shared ReadOnly Property Guam As CountryCodeInfo
            Get
                Return _Guam
            End Get
        End Property

        Public Shared ReadOnly Property Guatemala As CountryCodeInfo
            Get
                Return _Guatemala
            End Get
        End Property

        Public Shared ReadOnly Property Guernsey As CountryCodeInfo
            Get
                Return _Guernsey
            End Get
        End Property

        Public Shared ReadOnly Property Guinea As CountryCodeInfo
            Get
                Return _Guinea
            End Get
        End Property

        Public Shared ReadOnly Property GuineaBissau As CountryCodeInfo
            Get
                Return _GuineaBissau
            End Get
        End Property

        Public Shared ReadOnly Property Guyana As CountryCodeInfo
            Get
                Return _Guyana
            End Get
        End Property

        Public Shared ReadOnly Property Haiti As CountryCodeInfo
            Get
                Return _Haiti
            End Get
        End Property

        Public Shared ReadOnly Property Honduras As CountryCodeInfo
            Get
                Return _Honduras
            End Get
        End Property

        Public Shared ReadOnly Property HongKong As CountryCodeInfo
            Get
                Return _HongKong
            End Get
        End Property

        Public Shared ReadOnly Property Hungary As CountryCodeInfo
            Get
                Return _Hungary
            End Get
        End Property

        Public Shared ReadOnly Property Iceland As CountryCodeInfo
            Get
                Return _Iceland
            End Get
        End Property

        Public Shared ReadOnly Property ICOGlobal As CountryCodeInfo
            Get
                Return _ICOGlobal
            End Get
        End Property

        Public Shared ReadOnly Property India As CountryCodeInfo
            Get
                Return _India
            End Get
        End Property

        Public Shared ReadOnly Property Indonesia As CountryCodeInfo
            Get
                Return _Indonesia
            End Get
        End Property

        Public Shared ReadOnly Property InmarsatSNAC As CountryCodeInfo
            Get
                Return _InmarsatSNAC
            End Get
        End Property

        Public Shared ReadOnly Property InternationalFreephoneService As CountryCodeInfo
            Get
                Return _InternationalFreephoneService
            End Get
        End Property

        Public Shared ReadOnly Property InternationalSharedCostService As CountryCodeInfo
            Get
                Return _InternationalSharedCostService
            End Get
        End Property

        Public Shared ReadOnly Property Iran As CountryCodeInfo
            Get
                Return _Iran
            End Get
        End Property

        Public Shared ReadOnly Property Iraq As CountryCodeInfo
            Get
                Return _Iraq
            End Get
        End Property

        Public Shared ReadOnly Property Ireland As CountryCodeInfo
            Get
                Return _Ireland
            End Get
        End Property

        Public Shared ReadOnly Property Iridium As CountryCodeInfo
            Get
                Return _Iridium
            End Get
        End Property

        Public Shared ReadOnly Property IsleofMan As CountryCodeInfo
            Get
                Return _IsleofMan
            End Get
        End Property

        Public Shared ReadOnly Property Israel As CountryCodeInfo
            Get
                Return _Israel
            End Get
        End Property

        Public Shared ReadOnly Property Italy As CountryCodeInfo
            Get
                Return _Italy
            End Get
        End Property

        Public Shared ReadOnly Property Jamaica As CountryCodeInfo
            Get
                Return _Jamaica
            End Get
        End Property

        Public Shared ReadOnly Property JanMayen As CountryCodeInfo
            Get
                Return _JanMayen
            End Get
        End Property

        Public Shared ReadOnly Property Japan As CountryCodeInfo
            Get
                Return _Japan
            End Get
        End Property

        Public Shared ReadOnly Property Jersey As CountryCodeInfo
            Get
                Return _Jersey
            End Get
        End Property

        Public Shared ReadOnly Property Jordan As CountryCodeInfo
            Get
                Return _Jordan
            End Get
        End Property

        Public Shared ReadOnly Property Kazakhstan As CountryCodeInfo
            Get
                Return _Kazakhstan
            End Get
        End Property

        Public Shared ReadOnly Property Kenya As CountryCodeInfo
            Get
                Return _Kenya
            End Get
        End Property

        Public Shared ReadOnly Property Kiribati As CountryCodeInfo
            Get
                Return _Kiribati
            End Get
        End Property

        Public Shared ReadOnly Property NorthKorea As CountryCodeInfo
            Get
                Return _NorthKorea
            End Get
        End Property

        Public Shared ReadOnly Property SouthKorea As CountryCodeInfo
            Get
                Return _SouthKorea
            End Get
        End Property

        Public Shared ReadOnly Property Kuwait As CountryCodeInfo
            Get
                Return _Kuwait
            End Get
        End Property

        Public Shared ReadOnly Property Kyrgyzstan As CountryCodeInfo
            Get
                Return _Kyrgyzstan
            End Get
        End Property

        Public Shared ReadOnly Property Laos As CountryCodeInfo
            Get
                Return _Laos
            End Get
        End Property

        Public Shared ReadOnly Property Latvia As CountryCodeInfo
            Get
                Return _Latvia
            End Get
        End Property

        Public Shared ReadOnly Property Lebanon As CountryCodeInfo
            Get
                Return _Lebanon
            End Get
        End Property

        Public Shared ReadOnly Property Lesotho As CountryCodeInfo
            Get
                Return _Lesotho
            End Get
        End Property

        Public Shared ReadOnly Property Liberia As CountryCodeInfo
            Get
                Return _Liberia
            End Get
        End Property

        Public Shared ReadOnly Property Libya As CountryCodeInfo
            Get
                Return _Libya
            End Get
        End Property

        Public Shared ReadOnly Property Liechtenstein As CountryCodeInfo
            Get
                Return _Liechtenstein
            End Get
        End Property

        Public Shared ReadOnly Property Lithuania As CountryCodeInfo
            Get
                Return _Lithuania
            End Get
        End Property

        Public Shared ReadOnly Property Luxembourg As CountryCodeInfo
            Get
                Return _Luxembourg
            End Get
        End Property

        Public Shared ReadOnly Property Macau As CountryCodeInfo
            Get
                Return _Macau
            End Get
        End Property

        Public Shared ReadOnly Property Macedonia As CountryCodeInfo
            Get
                Return _Macedonia
            End Get
        End Property

        Public Shared ReadOnly Property Madagascar As CountryCodeInfo
            Get
                Return _Madagascar
            End Get
        End Property

        Public Shared ReadOnly Property Malawi As CountryCodeInfo
            Get
                Return _Malawi
            End Get
        End Property

        Public Shared ReadOnly Property Malaysia As CountryCodeInfo
            Get
                Return _Malaysia
            End Get
        End Property

        Public Shared ReadOnly Property Maldives As CountryCodeInfo
            Get
                Return _Maldives
            End Get
        End Property

        Public Shared ReadOnly Property Mali As CountryCodeInfo
            Get
                Return _Mali
            End Get
        End Property

        Public Shared ReadOnly Property Malta As CountryCodeInfo
            Get
                Return _Malta
            End Get
        End Property

        Public Shared ReadOnly Property MarshallIslands As CountryCodeInfo
            Get
                Return _MarshallIslands
            End Get
        End Property

        Public Shared ReadOnly Property Martinique As CountryCodeInfo
            Get
                Return _Martinique
            End Get
        End Property

        Public Shared ReadOnly Property Mauritania As CountryCodeInfo
            Get
                Return _Mauritania
            End Get
        End Property

        Public Shared ReadOnly Property Mauritius As CountryCodeInfo
            Get
                Return _Mauritius
            End Get
        End Property

        Public Shared ReadOnly Property Mayotte As CountryCodeInfo
            Get
                Return _Mayotte
            End Get
        End Property

        Public Shared ReadOnly Property Mexico As CountryCodeInfo
            Get
                Return _Mexico
            End Get
        End Property

        Public Shared ReadOnly Property FederatedStatesofMicronesia As CountryCodeInfo
            Get
                Return _FederatedStatesofMicronesia
            End Get
        End Property

        Public Shared ReadOnly Property MidwayIsland As CountryCodeInfo
            Get
                Return _MidwayIsland
            End Get
        End Property

        Public Shared ReadOnly Property Moldova As CountryCodeInfo
            Get
                Return _Moldova
            End Get
        End Property

        Public Shared ReadOnly Property Monaco As CountryCodeInfo
            Get
                Return _Monaco
            End Get
        End Property

        Public Shared ReadOnly Property Mongolia As CountryCodeInfo
            Get
                Return _Mongolia
            End Get
        End Property

        Public Shared ReadOnly Property Montenegro As CountryCodeInfo
            Get
                Return _Montenegro
            End Get
        End Property

        Public Shared ReadOnly Property Montserrat As CountryCodeInfo
            Get
                Return _Montserrat
            End Get
        End Property

        Public Shared ReadOnly Property Morocco As CountryCodeInfo
            Get
                Return _Morocco
            End Get
        End Property

        Public Shared ReadOnly Property Mozambique As CountryCodeInfo
            Get
                Return _Mozambique
            End Get
        End Property

        Public Shared ReadOnly Property Namibia As CountryCodeInfo
            Get
                Return _Namibia
            End Get
        End Property

        Public Shared ReadOnly Property Nauru As CountryCodeInfo
            Get
                Return _Nauru
            End Get
        End Property

        Public Shared ReadOnly Property Nepal As CountryCodeInfo
            Get
                Return _Nepal
            End Get
        End Property

        Public Shared ReadOnly Property Netherlands As CountryCodeInfo
            Get
                Return _Netherlands
            End Get
        End Property

        Public Shared ReadOnly Property Nevis As CountryCodeInfo
            Get
                Return _Nevis
            End Get
        End Property

        Public Shared ReadOnly Property NewCaledonia As CountryCodeInfo
            Get
                Return _NewCaledonia
            End Get
        End Property

        Public Shared ReadOnly Property NewZealand As CountryCodeInfo
            Get
                Return _NewZealand
            End Get
        End Property

        Public Shared ReadOnly Property Nicaragua As CountryCodeInfo
            Get
                Return _Nicaragua
            End Get
        End Property

        Public Shared ReadOnly Property Niger As CountryCodeInfo
            Get
                Return _Niger
            End Get
        End Property

        Public Shared ReadOnly Property Nigeria As CountryCodeInfo
            Get
                Return _Nigeria
            End Get
        End Property

        Public Shared ReadOnly Property Niue As CountryCodeInfo
            Get
                Return _Niue
            End Get
        End Property

        Public Shared ReadOnly Property NorfolkIsland As CountryCodeInfo
            Get
                Return _NorfolkIsland
            End Get
        End Property

        Public Shared ReadOnly Property NorthernMarianaIslands As CountryCodeInfo
            Get
                Return _NorthernMarianaIslands
            End Get
        End Property

        Public Shared ReadOnly Property Norway As CountryCodeInfo
            Get
                Return _Norway
            End Get
        End Property

        Public Shared ReadOnly Property Oman As CountryCodeInfo
            Get
                Return _Oman
            End Get
        End Property

        Public Shared ReadOnly Property Pakistan As CountryCodeInfo
            Get
                Return _Pakistan
            End Get
        End Property

        Public Shared ReadOnly Property Palau As CountryCodeInfo
            Get
                Return _Palau
            End Get
        End Property

        Public Shared ReadOnly Property PalestinianTerritories As CountryCodeInfo
            Get
                Return _PalestinianTerritories
            End Get
        End Property

        Public Shared ReadOnly Property Panama As CountryCodeInfo
            Get
                Return _Panama
            End Get
        End Property

        Public Shared ReadOnly Property PapuaNewGuinea As CountryCodeInfo
            Get
                Return _PapuaNewGuinea
            End Get
        End Property

        Public Shared ReadOnly Property Paraguay As CountryCodeInfo
            Get
                Return _Paraguay
            End Get
        End Property

        Public Shared ReadOnly Property Peru As CountryCodeInfo
            Get
                Return _Peru
            End Get
        End Property

        Public Shared ReadOnly Property Philippines As CountryCodeInfo
            Get
                Return _Philippines
            End Get
        End Property

        Public Shared ReadOnly Property PitcairnIslands As CountryCodeInfo
            Get
                Return _PitcairnIslands
            End Get
        End Property

        Public Shared ReadOnly Property Poland As CountryCodeInfo
            Get
                Return _Poland
            End Get
        End Property

        Public Shared ReadOnly Property Portugal As CountryCodeInfo
            Get
                Return _Portugal
            End Get
        End Property

        Public Shared ReadOnly Property PuertoRico As CountryCodeInfo
            Get
                Return _PuertoRico
            End Get
        End Property

        Public Shared ReadOnly Property Qatar As CountryCodeInfo
            Get
                Return _Qatar
            End Get
        End Property

        Public Shared ReadOnly Property Réunion As CountryCodeInfo
            Get
                Return _Réunion
            End Get
        End Property

        Public Shared ReadOnly Property Romania As CountryCodeInfo
            Get
                Return _Romania
            End Get
        End Property

        Public Shared ReadOnly Property Russia As CountryCodeInfo
            Get
                Return _Russia
            End Get
        End Property

        Public Shared ReadOnly Property Rwanda As CountryCodeInfo
            Get
                Return _Rwanda
            End Get
        End Property

        Public Shared ReadOnly Property Saba As CountryCodeInfo
            Get
                Return _Saba
            End Get
        End Property

        Public Shared ReadOnly Property SaintBarthélemy As CountryCodeInfo
            Get
                Return _SaintBarthélemy
            End Get
        End Property

        Public Shared ReadOnly Property SaintHelena As CountryCodeInfo
            Get
                Return _SaintHelena
            End Get
        End Property

        Public Shared ReadOnly Property SaintKittsandNevis As CountryCodeInfo
            Get
                Return _SaintKittsandNevis
            End Get
        End Property

        Public Shared ReadOnly Property SaintLucia As CountryCodeInfo
            Get
                Return _SaintLucia
            End Get
        End Property

        Public Shared ReadOnly Property SaintMartin As CountryCodeInfo
            Get
                Return _SaintMartin
            End Get
        End Property

        Public Shared ReadOnly Property SaintPierreandMiquelon As CountryCodeInfo
            Get
                Return _SaintPierreandMiquelon
            End Get
        End Property

        Public Shared ReadOnly Property SaintVincentandtheGrenadines As CountryCodeInfo
            Get
                Return _SaintVincentandtheGrenadines
            End Get
        End Property

        Public Shared ReadOnly Property Samoa As CountryCodeInfo
            Get
                Return _Samoa
            End Get
        End Property

        Public Shared ReadOnly Property SanMarino As CountryCodeInfo
            Get
                Return _SanMarino
            End Get
        End Property

        Public Shared ReadOnly Property SãoToméandPríncipe As CountryCodeInfo
            Get
                Return _SãoToméandPríncipe
            End Get
        End Property

        Public Shared ReadOnly Property SaudiArabia As CountryCodeInfo
            Get
                Return _SaudiArabia
            End Get
        End Property

        Public Shared ReadOnly Property Senegal As CountryCodeInfo
            Get
                Return _Senegal
            End Get
        End Property

        Public Shared ReadOnly Property Serbia As CountryCodeInfo
            Get
                Return _Serbia
            End Get
        End Property

        Public Shared ReadOnly Property Seychelles As CountryCodeInfo
            Get
                Return _Seychelles
            End Get
        End Property

        Public Shared ReadOnly Property SierraLeone As CountryCodeInfo
            Get
                Return _SierraLeone
            End Get
        End Property

        Public Shared ReadOnly Property Singapore As CountryCodeInfo
            Get
                Return _Singapore
            End Get
        End Property

        Public Shared ReadOnly Property SintEustatius As CountryCodeInfo
            Get
                Return _SintEustatius
            End Get
        End Property

        Public Shared ReadOnly Property SintMaarten As CountryCodeInfo
            Get
                Return _SintMaarten
            End Get
        End Property

        Public Shared ReadOnly Property Slovakia As CountryCodeInfo
            Get
                Return _Slovakia
            End Get
        End Property

        Public Shared ReadOnly Property Slovenia As CountryCodeInfo
            Get
                Return _Slovenia
            End Get
        End Property

        Public Shared ReadOnly Property SolomonIslands As CountryCodeInfo
            Get
                Return _SolomonIslands
            End Get
        End Property

        Public Shared ReadOnly Property Somalia As CountryCodeInfo
            Get
                Return _Somalia
            End Get
        End Property

        Public Shared ReadOnly Property SouthAfrica As CountryCodeInfo
            Get
                Return _SouthAfrica
            End Get
        End Property

        Public Shared ReadOnly Property SouthGeorgiaandtheSouthSandwichIslands As CountryCodeInfo
            Get
                Return _SouthGeorgiaandtheSouthSandwichIslands
            End Get
        End Property

        Public Shared ReadOnly Property SouthOssetia As CountryCodeInfo
            Get
                Return _SouthOssetia
            End Get
        End Property

        Public Shared ReadOnly Property SouthSudan As CountryCodeInfo
            Get
                Return _SouthSudan
            End Get
        End Property

        Public Shared ReadOnly Property Spain As CountryCodeInfo
            Get
                Return _Spain
            End Get
        End Property

        Public Shared ReadOnly Property SriLanka As CountryCodeInfo
            Get
                Return _SriLanka
            End Get
        End Property

        Public Shared ReadOnly Property Sudan As CountryCodeInfo
            Get
                Return _Sudan
            End Get
        End Property

        Public Shared ReadOnly Property Suriname As CountryCodeInfo
            Get
                Return _Suriname
            End Get
        End Property

        Public Shared ReadOnly Property Svalbard As CountryCodeInfo
            Get
                Return _Svalbard
            End Get
        End Property

        Public Shared ReadOnly Property Swaziland As CountryCodeInfo
            Get
                Return _Swaziland
            End Get
        End Property

        Public Shared ReadOnly Property Sweden As CountryCodeInfo
            Get
                Return _Sweden
            End Get
        End Property

        Public Shared ReadOnly Property Switzerland As CountryCodeInfo
            Get
                Return _Switzerland
            End Get
        End Property

        Public Shared ReadOnly Property Syria As CountryCodeInfo
            Get
                Return _Syria
            End Get
        End Property

        Public Shared ReadOnly Property Taiwan As CountryCodeInfo
            Get
                Return _Taiwan
            End Get
        End Property

        Public Shared ReadOnly Property Tajikistan As CountryCodeInfo
            Get
                Return _Tajikistan
            End Get
        End Property

        Public Shared ReadOnly Property Tanzania As CountryCodeInfo
            Get
                Return _Tanzania
            End Get
        End Property

        Public Shared ReadOnly Property Thailand As CountryCodeInfo
            Get
                Return _Thailand
            End Get
        End Property

        Public Shared ReadOnly Property Thuraya As CountryCodeInfo
            Get
                Return _Thuraya
            End Get
        End Property

        Public Shared ReadOnly Property Togo As CountryCodeInfo
            Get
                Return _Togo
            End Get
        End Property

        Public Shared ReadOnly Property Tokelau As CountryCodeInfo
            Get
                Return _Tokelau
            End Get
        End Property

        Public Shared ReadOnly Property Tonga As CountryCodeInfo
            Get
                Return _Tonga
            End Get
        End Property

        Public Shared ReadOnly Property TrinidadandTobago As CountryCodeInfo
            Get
                Return _TrinidadandTobago
            End Get
        End Property

        Public Shared ReadOnly Property TristandaCunha As CountryCodeInfo
            Get
                Return _TristandaCunha
            End Get
        End Property

        Public Shared ReadOnly Property Tunisia As CountryCodeInfo
            Get
                Return _Tunisia
            End Get
        End Property

        Public Shared ReadOnly Property Turkey As CountryCodeInfo
            Get
                Return _Turkey
            End Get
        End Property

        Public Shared ReadOnly Property Turkmenistan As CountryCodeInfo
            Get
                Return _Turkmenistan
            End Get
        End Property

        Public Shared ReadOnly Property TurksandCaicosIslands As CountryCodeInfo
            Get
                Return _TurksandCaicosIslands
            End Get
        End Property

        Public Shared ReadOnly Property Tuvalu As CountryCodeInfo
            Get
                Return _Tuvalu
            End Get
        End Property

        Public Shared ReadOnly Property Uganda As CountryCodeInfo
            Get
                Return _Uganda
            End Get
        End Property

        Public Shared ReadOnly Property Ukraine As CountryCodeInfo
            Get
                Return _Ukraine
            End Get
        End Property

        Public Shared ReadOnly Property UnitedArabEmirates As CountryCodeInfo
            Get
                Return _UnitedArabEmirates
            End Get
        End Property

        Public Shared ReadOnly Property UnitedKingdom As CountryCodeInfo
            Get
                Return _UnitedKingdom
            End Get
        End Property

        Public Shared ReadOnly Property UnitedStates As CountryCodeInfo
            Get
                Return _UnitedStates
            End Get
        End Property

        Public Shared ReadOnly Property UniversalPersonalTelecommunications As CountryCodeInfo
            Get
                Return _UniversalPersonalTelecommunications
            End Get
        End Property

        Public Shared ReadOnly Property Uruguay As CountryCodeInfo
            Get
                Return _Uruguay
            End Get
        End Property

        Public Shared ReadOnly Property USVirginIslands As CountryCodeInfo
            Get
                Return _USVirginIslands
            End Get
        End Property

        Public Shared ReadOnly Property Uzbekistan As CountryCodeInfo
            Get
                Return _Uzbekistan
            End Get
        End Property

        Public Shared ReadOnly Property Vanuatu As CountryCodeInfo
            Get
                Return _Vanuatu
            End Get
        End Property

        Public Shared ReadOnly Property Venezuela As CountryCodeInfo
            Get
                Return _Venezuela
            End Get
        End Property

        Public Shared ReadOnly Property VaticanCityState As CountryCodeInfo
            Get
                Return _VaticanCityState
            End Get
        End Property

        Public Shared ReadOnly Property Vietnam As CountryCodeInfo
            Get
                Return _Vietnam
            End Get
        End Property

        Public Shared ReadOnly Property WakeIsland As CountryCodeInfo
            Get
                Return _WakeIsland
            End Get
        End Property

        Public Shared ReadOnly Property WallisandFutuna As CountryCodeInfo
            Get
                Return _WallisandFutuna
            End Get
        End Property

        Public Shared ReadOnly Property Yemen As CountryCodeInfo
            Get
                Return _Yemen
            End Get
        End Property

        Public Shared ReadOnly Property Zambia As CountryCodeInfo
            Get
                Return _Zambia
            End Get
        End Property

        Public Shared ReadOnly Property Zanzibar As CountryCodeInfo
            Get
                Return _Zanzibar
            End Get
        End Property

        Public Shared ReadOnly Property Zimbabwe As CountryCodeInfo
            Get
                Return _Zimbabwe
            End Get
        End Property

#End Region

        ''' <summary>
        ''' Name of the country or service.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Name As String
            Get
                Return _Name
            End Get
        End Property

        ''' <summary>
        ''' Indicates whether this entry is for a satellite service, and not a country
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property IsSatelliteService As Boolean
            Get
                Return _IsSatService
            End Get
        End Property

        ''' <summary>
        ''' Returns a list of full calling codes for a specific country
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Codes As List(Of String)
            Get
                Return _Codes
            End Get
        End Property

        ''' <summary>
        ''' Returns the default full calling code for a specific country
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Code As String
            Get
                Return _Codes(0)
            End Get
        End Property

        ''' <summary>
        ''' Returns the default primary code for a specific country
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property PrimaryCode As String
            Get
                Return _PrimaryCodes(0)
            End Get
        End Property

        ''' <summary>
        ''' Returns a list of all primary codes valid for the specific country
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property PrimaryCodes As List(Of String)
            Get
                Return _PrimaryCodes
            End Get
        End Property

        ''' <summary>
        ''' Returns the main country associated with the first primary code of this instance.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property MainCountry As CountryCodeInfo
            Get
                If PrimaryCode = Code Then Return Me
                Dim cc As CountryCodeInfo = FindMainCountry(PrimaryCode)

                Return cc
            End Get
        End Property

        ''' <summary>
        ''' Boolean value indicating whether the country is a member of the North American Numbering Plan (NANP)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property NANP As Boolean
            Get
                Return ((Code = "+1") Or (PrimaryCode = "+1"))
            End Get
        End Property

        ''' <summary>
        ''' Parse a code for a primary and return the primary and full country codes.
        ''' </summary>
        ''' <param name="code">Raw string to parse</param>
        ''' <param name="primary">Primary country code returned.</param>
        ''' <param name="fullCode">Full country code (including primary part).</param>
        ''' <remarks></remarks>
        Private Shared Sub ParseCode(code As String, ByRef primary As String, ByRef fullCode As String)
            '    code = SearchReplace(code, "+", "")
            Dim i As Integer

            i = code.IndexOf(" ")
            If i >= 0 Then
                primary = code.Substring(0, i)
            Else
                primary = ""
            End If

            fullCode = code
        End Sub

        ''' <summary>
        ''' Attempts to find a country or satellite service matching the specific code.
        ''' </summary>
        ''' <param name="code">Country code to match.</param>
        ''' <returns>A list of all matched countries.</returns>
        ''' <remarks></remarks>
        Public Shared Function Find(code As String) As List(Of CountryCodeInfo)
            Dim tf As String = NoSpace(code).Replace("+", "")
            Dim nl As New List(Of CountryCodeInfo)

            For Each cc In _AllCountries
                For Each cde In cc.Codes
                    cde = NoSpace(cde).Replace("+", "")
                    If cde = tf Then
                        If Not nl.Contains(cc) Then nl.Add(cc)
                        Exit For
                    End If
                Next

                For Each cde In cc.PrimaryCodes
                    cde = NoSpace(cde).Replace("+", "")
                    If cde = tf Then
                        If Not nl.Contains(cc) Then nl.Add(cc)
                        Exit For
                    End If
                Next
            Next

            For Each cc In _AllSat
                For Each cde In cc.Codes
                    cde = NoSpace(cde).Replace("+", "")
                    If cde = tf Then
                        If Not nl.Contains(cc) Then nl.Add(cc)
                        Exit For
                    End If
                Next

                For Each cde In cc.PrimaryCodes
                    cde = NoSpace(cde).Replace("+", "")
                    If cde = tf Then
                        If Not nl.Contains(cc) Then nl.Add(cc)
                        Exit For
                    End If
                Next
            Next

            Return nl
        End Function

        ''' <summary>
        ''' Find the main country to which the primary code belongs.
        ''' </summary>
        ''' <param name="primary">Primary country code to search (e.g. "+1")</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FindMainCountry(primary As String) As CountryCodeInfo
            For Each cc In _AllCountries
                If cc.Code = primary Then Return cc
            Next

            Return Nothing
        End Function

        ''' <summary>
        ''' Returns a list of all countries in the North American Numbering Plan (NANP)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property NANPCountries As List(Of CountryCodeInfo)
            Get
                Return Find("+1")
            End Get
        End Property

        ''' <summary>
        ''' Returns a list of all country codes that do not have secondary modifiers.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property AllPrimaryCountries As List(Of CountryCodeInfo)
            Get
                Dim nl As New List(Of CountryCodeInfo)

                For Each cc In _AllCountries
                    If cc.PrimaryCode = "" Then nl.Add(cc)
                Next

                Return nl
            End Get
        End Property

        ''' <summary>
        ''' Return a list of all countries that contain the same primary country code.
        ''' </summary>
        ''' <param name="primary">Primary country code to search (e.g. "+1")</param>
        ''' <param name="includeMain">Specify whether to include the main country in the returned search.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AllCountriesInPrimary(primary As String, Optional includeMain As Boolean = True) As List(Of CountryCodeInfo)
            Dim nl As New List(Of CountryCodeInfo)
            For Each cc In _AllCountries

                If includeMain Then
                    For Each cde In cc.Codes
                        If cde = primary Then
                            If Not nl.Contains(cc) Then nl.Add(cc)
                            Exit For
                        End If
                    Next
                End If

                For Each cde In cc.PrimaryCodes
                    If cde = primary Then
                        If Not nl.Contains(cc) Then nl.Add(cc)
                        Exit For
                    End If
                Next

            Next

            Return nl
        End Function

        ''' <summary>
        ''' Return a list of all country codes in the present execution context.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property AllCountryCodes As List(Of CountryCodeInfo)
            Get
                Return _AllCountries
            End Get
        End Property

        ''' <summary>
        ''' Return a list of all satellite-service country codes in the present execution context.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property AllSatelliteServiceCodes As List(Of CountryCodeInfo)
            Get
                Return _AllSat
            End Get
        End Property

        ''' <summary>
        ''' Returns a string representation of this object.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function ToString() As String
            Return Code & ": " & Name
        End Function

        ''' <summary>
        ''' Instantiate a new CountryCodeInfo class.
        ''' </summary>
        ''' <param name="sat">Boolean value that specifies whether or not this is a country code for a global satellite phone service.</param>
        ''' <param name="country">Name of country or organization.</param>
        ''' <param name="codes">List of country codes</param>
        ''' <remarks></remarks>
        Public Sub New(sat As Boolean, country As String, ParamArray codes() As String)
            _IsSatService = sat
            _Name = country
            For Each cc In codes
                Dim pc As String = "",
                co As String = ""

                ParseCode(cc, pc, co)
                _Codes.Add(co)
                _PrimaryCodes.Add(pc)
            Next

            If sat Then _AllSat.Add(Me) Else _AllCountries.Add(Me)
        End Sub

    End Class

End Namespace