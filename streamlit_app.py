for r in range(ponynamen_start_index, max_rij):
    if eigen_pony_rij is not None and r >= eigen_pony_rij:
        break  # Stop zodra we bij of voorbij "eigen pony" zijn

    naam = str(df.iloc[r, col])
    pony = str(df.iloc[r, ponynamen_kolom])
    if not naam.strip() or naam.strip().lower() in ["", "nan", "x"]:
        continue

    delen = naam.strip().split()
    voornaam = delen[0].capitalize() if delen else ""
    achternaam = ""
    tussenvoegsels = {"van", "de", "der", "den", "ter", "ten", "het", "te"}
    for deel in delen[1:]:
        if deel.lower() not in tussenvoegsels:
            achternaam = deel.capitalize()
            break
    code = voornaam
    key = voornaam.lower()
    if key in namen_counter:
        code += achternaam[:1].upper()
    namen_counter[key] = namen_counter.get(key, 0) + 1

    opmerking = ""
    for sleutel, tekst in st.session_state.pony_opmerkingen.items():
        if sleutel.lower() in pony.lower():
            opmerking = f" ({tekst})"
            break

    locatie = "(B)" if pony in reeds_in_bak else "(S)"
    pony_tekst = f"{pony.title()} {locatie}{opmerking}"
    kind_pony_combinaties.append((code, pony_tekst))
    reeds_in_bak.add(pony)
