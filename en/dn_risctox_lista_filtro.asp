<%
	filtro=EliminaInyeccionSQL(request("f"))
	select case filtro
		case "negra":
			titulo_lista="RISCTOX: Substances of concern for Trade unions"
		case "cym":
			titulo_lista="RISCTOX: Carcinogens and mutagens > According to Regulation 1272/2008"
		case "cym2":
			titulo_lista="RISCTOX: Carcinogens and mutagens > According to IARC"
		case "cym3":
			titulo_lista="RISCTOX: Carcinogens and mutagens > According to other sources"
		case "mama":
			titulo_lista="RISCTOX: Carcinogens and mutagens > According to SSI (breast cancer)"
		case "tpr":
			titulo_lista="RISCTOX: Reproductive toxicans"
		case "dis":
			titulo_lista="RISCTOX: Endocrine disrupters"
		case "neu":
			titulo_lista="RISCTOX: Neurotoxicants"
		case "oto":
			titulo_lista="RISCTOX: Neurotoxicants > Ototoxicans"
		case "sen":
			titulo_lista="RISCTOX: Sensitisers"
		case "senreach":
			titulo_lista="RISCTOX: Sensitisers > REACH allergens"
		case "pyb":
			titulo_lista="RISCTOX: Persistent, Bioaccumulative and Toxics"
		case "mpmb":
			titulo_lista="RISCTOX: vPvB"
		case "tac":
			titulo_lista="RISCTOX: Aquatic toxicity > Water Frame Directive"
		case "tac2":
			titulo_lista="RISCTOX: Aquatic toxicity > German water pollutants"
		case "dat":
			titulo_lista="RISCTOX: Atmospheric pollutants > Ozone-depleting substances"
		case "dat2":
			titulo_lista="RISCTOX: Atmospheric pollutants > Greenhouse gases"
		case "dat3":
			titulo_lista="RISCTOX: Atmospheric pollutants > Air pollutants"
		case "cos":
			titulo_lista="RISCTOX: Soil pollutants > Regulation 9/2005"
		case "cop":
			titulo_lista="RISCTOX: Persistent Organic Pollutants (POPs)"
		case "enf":
			titulo_lista="RISCTOX: Occupational Health and Safety Regulations > Substance linked with Occupational diseases"
		case "cov":
			titulo_lista="RISCTOX: Environmental regulations > VOCs"
		case "ep1":
			titulo_lista="RISCTOX: Environmental regulations > IPPC > PRTR (Water)"
		case "ep2":
			titulo_lista="RISCTOX: Environmental regulations > IPPC > PRTR (Air)"
		case "ep3":
			titulo_lista="RISCTOX: Environmental regulations > IPPC > PRTR (Soil)"
		case "pro":
			titulo_lista="RISCTOX: Regulations on restriction / prohibition of substances > Banned substances"
		case "rest":
			titulo_lista="RISCTOX: Regulations on restriction / prohibition of substances > Restricted substances under REACH"
		case "candidatas_reach":
			titulo_lista="RISCTOX: Regulations on restriction / prohibition of substances > REACH Candidate list"
		case "autorizacion_reach":
			titulo_lista="RISCTOX: Regulations on restriction / prohibition of substances > REACH Authorisation list"
		case "biocidas_prohibidas":
			titulo_lista="RISCTOX: Regulations on restriction / prohibition of substances > Banned biocides"
		case "biocidas_autorizadas":
			titulo_lista="RISCTOX: Regulations on restriction / prohibition of substances > Authorised biocides"
		case "pesticidas_autorizadas":
			titulo_lista="RISCTOX: Regulations on restriction / prohibition of substances > Authorised pesticides"
		case "pesticidas_prohibidas":
			titulo_lista="RISCTOX: Regulations on restriction / prohibition of substances > Banned pesticides"
		case "corap"
			titulo_lista="RISCTOX: Regulations on restriction / prohibition of substances > Substances under CORAP evaluation"
	end select
	form_action = "dn_risctox_lista.asp?busc=1&f="&filtro

%>
