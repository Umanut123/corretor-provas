import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import { createClient } from "npm:@supabase/supabase-js@2";
import { encodeBase64 } from "jsr:@std/encoding@1/base64";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { ...CORS, "Content-Type": "application/json" } });

const MODELO = Deno.env.get("MODELO") || "claude-opus-5-5";

async function claude(texto: string, imagens: { data: string; type: string }[], maxTokens = 8000) {
  const key = Deno.env.get("ANTHROPIC_API_KEY");
  if (!key) throw new Error("A chave da Anthropic (ANTHROPIC_API_KEY) não foi configurada no Supabase.");
  const content = [
    ...imagens.map((im) => ({ type: "image", source: { type: "base64", media_type: im.type, data: im.data } })),
    { type: "text", text: texto },
  ];
  const r = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json" },
    body: JSON.stringify({ model: MODELO, max_tokens: maxTokens, messages: [{ role: "user", content }] }),
  });
  const body = await r.json();
  if (!r.ok) {
    const msg = body?.error?.message || r.statusText;
    if (r.status === 401) throw new Error("Chave da Anthropic inválida.");
    if (r.status === 429) throw new Error("Limite de uso da IA atingido. Tente em alguns minutos.");
    if (/credit|billing/i.test(msg)) throw new Error("Sem créditos na conta da Anthropic.");
    throw new Error("Erro da IA: " + msg);
  }
  return (body.content || []).filter((c: { type: string }) => c.type === "text").map((c: { text: string }) => c.text).join("");
}
function parseJSON(t: string) {
  try { return JSON.parse(t); } catch { /* segue */ }
  const f = t.match(/```(?:json)?\s*([\s\S]*?)```/); if (f) { try { return JSON.parse(f[1]); } catch { /* segue */ } }
  const a = t.indexOf("{"), b = t.lastIndexOf("}");
  if (a >= 0 && b > a) return JSON.parse(t.slice(a, b + 1));
  throw new Error("A IA respondeu em formato inesperado. Tente de novo.");
}
async function baixarFotos(sb: ReturnType<typeof createClient>, paths: string[]) {
  const out: { data: string; type: string }[] = [];
  for (const p of paths) {
    const { data, error } = await sb.storage.from("provas").download(p);
    if (error || !data) throw new Error("Não consegui abrir uma das fotos.");
    out.push({ data: encodeBase64(new Uint8Array(await data.arrayBuffer())), type: data.type || "image/jpeg" });
  }
  return out;
}
const num = (v: unknown) => { const n = parseFloat(String(v ?? "").replace(",", ".")); return isNaN(n) ? null : n; };
const r2 = (n: number) => Math.round(n * 100) / 100;

function promptCorrecao(prova: any, disc: any, alunos: any[]) {
  return `Você está ajudando um professor a corrigir uma prova manuscrita. As imagens anexas são as páginas da prova respondida por UM aluno, em ordem.

DISCIPLINA: ${disc.nome}${disc.turma ? ` — turma ${disc.turma}` : ""}
PROVA: ${prova.titulo}

GABARITO DO PROFESSOR (respostas corretas e valor de cada questão):
"""
${prova.gabarito.trim()}
"""
${prova.modelo?.trim() ? `\nMODELO DA PROVA EM BRANCO (enunciados das questões, para você entender o que cada questão pede):\n"""\n${prova.modelo.trim().slice(0, 20000)}\n"""\n` : ""}${prova.instrucoes?.trim() ? `\nORIENTAÇÕES DO PROFESSOR: ${prova.instrucoes.trim()}\n` : ""}
ALUNOS CADASTRADOS NESTA DISCIPLINA (id: nome):
${alunos.length ? alunos.map((a) => `${a.id}: ${a.nome}`).join("\n") : "(nenhum ainda)"}

O QUE FAZER
1. Leia o nome do aluno escrito na prova (pode estar em hebraico, português ou misturado; escreva como está em "nome_lido").
   Se ele corresponder a um aluno da lista acima (mesmo escrito em outro alfabeto, com apelido ou abreviação), informe o id em "aluno_id"; senão null. Use "nome_incerto": true se não tiver certeza.
2. Para CADA questão do gabarito, na ordem do gabarito, localize a resposta do aluno e transcreva-a fielmente no idioma em que foi escrita (hebraico, português ou misturado — não traduza). Respostas longas: até 600 caracteres, corte com "[...]".
3. Avalie que porcentagem da resposta correta a resposta do aluno ATENDE, de 0 a 100, em múltiplos de 5:
   - 100 = atende plenamente; 0 = em branco, errada ou sem relação; intermediário conforme a parte do conteúdo esperado que foi contemplada.
   - Julgue o conteúdo e o sentido, não o idioma nem a ortografia.
   - Questões de alternativa: 0 ou 100.
4. "valor": quantos pontos a questão vale segundo o gabarito; se o gabarito não disser, null.
5. "obs": em português, 1–2 frases dizendo o que a resposta contemplou e o que faltou.
6. Se não conseguir ler um trecho com segurança, "legivel": false e explique em "obs". Nunca invente o que o aluno escreveu.
7. "comentario": devolutiva para o aluno, em português, falando com "você", 2 a 4 frases, respeitosa e encorajadora.

Responda SOMENTE com JSON:
{"nome_lido": "…", "aluno_id": "uuid ou null", "nome_incerto": false, "questoes": [{"q": "1", "valor": 10, "resposta": "…", "legivel": true, "pct": 80, "obs": "…"}], "comentario": "…"}`;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  try {
    const auth = req.headers.get("Authorization") || "";
    const url = Deno.env.get("SUPABASE_URL")!;
    const sb = createClient(url, Deno.env.get("SUPABASE_ANON_KEY")!, { global: { headers: { Authorization: auth } } });
    const admin = createClient(url, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);

    const { data: u } = await sb.auth.getUser();
    const email = u?.user?.email?.toLowerCase();
    if (!email) return json({ erro: "Faça login de novo." }, 401);
    const { data: ok } = await admin.from("usuarios_autorizados").select("email").eq("email", email).maybeSingle();
    if (!ok) return json({ erro: "Este usuário não está autorizado a usar a correção por IA." }, 403);

    const body = await req.json();

    // ---------- transcrever gabarito a partir de fotos ----------
    if (body.acao === "transcrever_gabarito") {
      const imgs = await baixarFotos(sb, body.fotos || []);
      const texto = await claude(`As imagens anexas são o GABARITO (respostas corretas) de uma prova, feito por um professor. Pode estar em português, hebraico ou misturado, digitado ou manuscrito.
Pode ser uma folha só de respostas ou a própria prova preenchida pelo professor com as respostas certas.
Transcreva com fidelidade, no idioma original (não traduza), mantendo a numeração das questões e o valor de cada questão quando aparecer (escreva como "(10 pontos)"). Para cada questão escreva a RESPOSTA CORRETA (se for alternativa, a letra e o texto dela).
Um item por questão, no formato:
1) (10 pontos) resposta correta…
2) …
Se algum trecho estiver ilegível, escreva [ilegível]. Responda só com a transcrição.`, imgs, 6000);
      return json({ texto: texto.trim() });
    }

    // ---------- transcrever modelo (prova em branco) ----------
    if (body.acao === "transcrever_modelo") {
      const imgs = await baixarFotos(sb, body.fotos || []);
      const texto = await claude(`As imagens anexas são o MODELO de uma prova em branco (as perguntas, sem respostas), feito por um professor. Pode estar em português, hebraico ou misturado.
Transcreva com fidelidade, no idioma original (não traduza), cada questão com sua numeração, o enunciado completo, as alternativas se houver e o valor/pontos se aparecer (escreva como "(10 pontos)").
Formato:
1) (10 pontos) enunciado…
   a) … b) …
2) …
Se algum trecho estiver ilegível, escreva [ilegível]. Responda só com a transcrição.`, imgs, 8000);
      return json({ texto: texto.trim() });
    }

    // ---------- corrigir a prova de um aluno ----------
    if (body.acao === "corrigir") {
      const { data: c, error: e1 } = await sb.from("correcoes").select("*").eq("id", body.correcao_id).single();
      if (e1 || !c) return json({ erro: "Correção não encontrada." }, 404);
      const { data: prova } = await sb.from("provas").select("*").eq("id", c.prova_id).single();
      const { data: disc } = await sb.from("disciplinas").select("*").eq("id", prova.disciplina_id).single();
      const { data: alunos } = await sb.from("alunos").select("id,nome").eq("disciplina_id", prova.disciplina_id).order("nome");
      if (!prova.gabarito?.trim()) return json({ erro: "Cadastre o gabarito da prova primeiro." }, 400);
      if (!c.fotos?.length) return json({ erro: "Esta prova não tem fotos." }, 400);

      await sb.from("correcoes").update({ status: "corrigindo", erro: null }).eq("id", c.id);
      try {
        const imgs = await baixarFotos(sb, c.fotos);
        const res = parseJSON(await claude(promptCorrecao(prova, disc, alunos || []), imgs));
        const padrao = Number(prova.valor_padrao) || 1;
        const questoes = (Array.isArray(res.questoes) ? res.questoes : []).map((q: any, i: number) => {
          const valor = (num(q?.valor) ?? 0) > 0 ? num(q.valor)! : padrao;
          const pct = Math.max(0, Math.min(100, Math.round(num(q?.pct) ?? 0)));
          return { q: String(q?.q ?? i + 1), valor, pct, nota: r2(valor * pct / 100), resposta: String(q?.resposta ?? ""), legivel: q?.legivel !== false, obs: String(q?.obs ?? "") };
        });
        if (!questoes.length) throw new Error("A IA não encontrou as questões. Confira o gabarito e as fotos.");
        const ids = new Set((alunos || []).map((a) => a.id));
        const alunoId = c.aluno_id || (res.aluno_id && ids.has(res.aluno_id) ? res.aluno_id : null);
        const upd = {
          status: "corrigido", erro: null, questoes,
          nota: r2(questoes.reduce((s: number, q: any) => s + q.nota, 0)),
          nota_max: r2(questoes.reduce((s: number, q: any) => s + q.valor, 0)),
          comentario: String(res.comentario || ""), nome_lido: res.nome_lido ? String(res.nome_lido) : null,
          nome_incerto: !!res.nome_incerto || (!alunoId && !!res.nome_lido), aluno_id: alunoId, revisado: false,
        };
        const { data: saved } = await sb.from("correcoes").update(upd).eq("id", c.id).select().single();
        return json({ correcao: saved });
      } catch (err) {
        await sb.from("correcoes").update({ status: "erro", erro: String((err as Error).message) }).eq("id", c.id);
        throw err;
      }
    }

    // ---------- análise da turma ----------
    if (body.acao === "analise") {
      const { data: prova } = await sb.from("provas").select("*").eq("id", body.prova_id).single();
      const { data: cs } = await sb.from("correcoes").select("questoes,nota,nota_max").eq("prova_id", body.prova_id).eq("status", "corrigido");
      const porQ: Record<string, { valor: number; notas: number[]; obs: string[] }> = {};
      for (const c of cs || []) for (const q of c.questoes || []) {
        (porQ[q.q] ||= { valor: q.valor, notas: [], obs: [] }).notas.push(q.nota);
        if (q.obs && porQ[q.q].obs.length < 15) porQ[q.q].obs.push(q.obs);
      }
      const dados = { prova: prova.titulo, gabarito: prova.gabarito.slice(0, 6000), alunos: cs?.length,
        questoes: Object.entries(porQ).map(([q, v]) => ({ q, valor: v.valor, media: r2(v.notas.reduce((s, x) => s + x, 0) / v.notas.length), observacoes: v.obs })) };
      const texto = await claude(`Você é coordenador pedagógico. Com base nos dados de correção abaixo, escreva para o professor, em português, uma análise curta da turma (até 250 palavras, texto simples sem markdown, parágrafos curtos): 1) visão geral; 2) questões/temas com mais dificuldade e o erro mais comum; 3) três sugestões práticas para a próxima aula. Não cite alunos.

DADOS:
${JSON.stringify(dados)}`, [], 2000);
      await sb.from("provas").update({ analise: texto.trim() }).eq("id", body.prova_id);
      return json({ texto: texto.trim() });
    }

    return json({ erro: "Ação desconhecida." }, 400);
  } catch (err) {
    console.error(err);
    return json({ erro: String((err as Error).message || err) }, 500);
  }
});
