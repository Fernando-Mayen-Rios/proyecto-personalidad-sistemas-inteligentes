-- ¿Cuántas personas hay de cada perfil?
SELECT * FROM v_resumen_perfiles;

-- ¿Cómo se comportan las variables según el perfil?
SELECT * FROM v_variables_por_perfil;

-- ¿Cómo se reparte el perfil por género?
SELECT genero, perfil, personas FROM v_perfil_por_genero WHERE perfil <> 'Por revisar';

-- ¿Qué modelo tuvo mejores resultados?
SELECT modelo, ROUND(aciertos_balanceados, 2) AS aciertos_balanceados, ROUND(f1, 2) AS f1, ROUND(auc, 2) AS auc FROM resultados_modelos ORDER BY aciertos_balanceados DESC;

-- ¿Qué personas No introvertidas son líderes o conversan con otros? (une dos tablas)
SELECT p.id_persona, p.perfil, m.rol_lider, m.act_conversar FROM personas p JOIN variables_modelo m ON m.id_persona = p.id_persona WHERE p.perfil = 'No introvertido' AND (m.rol_lider = 1 OR m.act_conversar = 1) ORDER BY p.id_persona;

