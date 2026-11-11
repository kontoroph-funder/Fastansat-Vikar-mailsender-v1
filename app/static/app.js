'use strict';
const $ = id => document.getElementById(id);
const months = ['januar','februar','marts','april','maj','juni','juli','august','september','oktober','november','december'];
const labels = {ANALYZING:'Analyserer',DRAFT:'Kladde',PAUSED:'Kræver kontrol',SENDING:'Sender',COMPLETE:'Afsluttet',CANCELLED:'Annulleret',EXPIRED:'Udløbet',ERROR:'Fejl',READY:'Klar',BLOCKED:'Fejl i ark',ACCEPTED:'Accepteret af Outlook',UNKNOWN:'Ukendt afsendelse',FAILED:'Afsendelse afvist'};
let csrf='', account=null, job=null, activeId=null, chosenFile=null, dirty=false, saving=null, timer=null, settings=null, busy=false, stopped=false;
let ccDirty=false, ccSaving=null, outlookAccounts=[];
const methodName = method => method==='outlook'?'Klassisk Outlook':'Microsoft Graph';
const statusLabel = status => status==='ACCEPTED'&&job?.mail_method==='outlook'?'Afleveret til Outlook':labels[status];
function showMethod(){const classic=$('mailMethod').value==='outlook';$('outlookSettings').classList.toggle('hidden',!classic);$('graphSettings').classList.toggle('hidden',classic);$('tenantId').required=!classic;}
function showOutlookAccounts(selected){$('outlookAccount').replaceChildren(new Option('Vælg en konto',''));for(const a of outlookAccounts)$('outlookAccount').append(new Option(a.email+' · '+a.name,a.id));$('outlookAccount').value=selected||'';}
const displayDate = value => new Intl.DateTimeFormat('da-DK',{dateStyle:'short',timeStyle:'short',timeZone:'Europe/Copenhagen'}).format(new Date(value*1000));
const row = () => job?.recipients.find(r=>r.id===activeId);
const canEdit = () => job && ['DRAFT','PAUSED'].includes(job.state) && job.test_status!=='SENDING' && job.expires>Date.now()/1000;
const locked = r => ['SENDING','ACCEPTED','UNKNOWN','FAILED'].includes(r.status);
const sendable = r => r.sendable && r.status==='READY' && !r.needs_review && !r.needs_pdf_review;
function text(tag, value, className){const el=document.createElement(tag);el.textContent=value;if(className)el.className=className;return el;}
function notice(value, info=false){$('notice').textContent=value;$('notice').className='notice'+(info?' info':'')+(value?'':' hidden');}
async function api(path, method='GET', body){
  const headers={'X-CSRF-Token':csrf};
  if(body && !(body instanceof FormData))headers['Content-Type']='application/json';
  const response=await fetch('/api'+path,{method,headers,body:body instanceof FormData?body:body?JSON.stringify(body):undefined});
  if(response.status===404&&path==='/outlook/accounts')throw new Error('Den kørende app er en ældre version. Klik Afslut appen, åbn appen igen fra genvejen, og genindlæs siden.');
  const result=await response.json();
  if(!response.ok)throw new Error(result.error || (Array.isArray(result.detail)?result.detail.map(e=>e.msg).join(' · '):result.detail) || 'Handlingen kunne ikke gennemføres.');
  return result;
}
async function action(fn){if(busy)return;busy=true;try{notice('');await fn();}catch(error){notice(error.message);}finally{busy=false;}}
function currentPeriod(){const m=Number($('month').value),y=Number($('year').value);return `16. ${months[(m+10)%12]}${m===1?' '+(y-1):''} – 15. ${months[m-1]} ${y}`;}
function setPeriod(m,y){if(!Array.from($('year').options).some(o=>Number(o.value)===y))$('year').append(new Option(y,y));$('month').value=m;$('year').value=y;$('periodLabel').textContent=currentPeriod();}
function showAccount(){const label=account?account.email+(settings?' · '+methodName(settings.mail_method):''):'Vælg afsenderkonto';$('accountLabel').textContent=label;$('connectionStatus').textContent=label;$('accountSettings').textContent=account?'Kontoindstillinger':'Forbind konto';}
async function tab(name){await saveDraft();document.querySelectorAll('.tab-panel').forEach(el=>el.classList.toggle('active',el.id===name+'Tab'));document.querySelectorAll('.nav-item[data-tab]').forEach(el=>el.classList.toggle('active',el.dataset.tab===name));if(name==='history')await history();}
async function saveDraft(){
  clearTimeout(timer);
  if(ccSaving)await ccSaving;
  if(ccDirty&&job){
    const cc=$('cc').value, jobId=job.id;
    $('ccStatus').textContent='Gemmer…';
    ccSaving=api(`/jobs/${jobId}/cc`,'PATCH',{cc});
    try{const updated=await ccSaving;if(job?.id===jobId){job=updated;if($('cc').value===cc){ccDirty=false;$('ccStatus').textContent='Gemt';}}}
    catch(e){$('ccStatus').textContent='Ikke gemt';throw e;}finally{ccSaving=null;}
  }
  if(saving)await saving;
  if(!dirty||!row())return;
  const rowId=activeId, jobId=job.id, subject=$('subject').value, body=$('body').value;
  $('saveStatus').textContent='Gemmer…';
  saving=api(`/jobs/${jobId}/recipients/${rowId}`,'PATCH',{subject,body});
  try{const updated=await saving;if(job?.id===jobId){job=updated;if(activeId===rowId&&$('subject').value===subject&&$('body').value===body){dirty=false;$('saveStatus').textContent='Gemt';}render(false);}}catch(e){$('saveStatus').textContent='Ikke gemt';throw e;}finally{saving=null;}
}
function render(refreshDetail=true){
  const rows=job?.recipients||[], ready=rows.filter(sendable), errors=rows.filter(r=>r.errors.length), selected=ready.filter(r=>r.selected);
  $('kind').disabled=!!job;$('kind').value=job?.kind||$('kind').value;
  const editable=canEdit();
  $('cc').disabled=!!job&&!editable;$('saveCc').disabled=!job||!editable;
  if(!ccDirty&&!ccSaving){$('cc').value=job?.cc||'';$('ccStatus').textContent=job?'Gemt':'';}
  $('month').disabled=$('year').disabled=!!job&&(!editable||rows.some(locked));
  $('chooseFile').disabled=$('analyze').disabled=!!job;
  $('fileName').textContent=job?.filename||chosenFile?.name||(job?.cleaned?'Arbejdsfiler slettet':'Ingen fil valgt');
  $('jobInfo').textContent=job?`${labels[job.state]} · ${methodName(job.mail_method)} · ${job.cleaned?'Indhold slettet':'Udløber '+displayDate(job.expires)}`:'Vælg en .xlsx-fil for at komme i gang.';
  $('statTotal').textContent=job?.cleaned?Object.values(job.counts||{}).reduce((a,b)=>a+b,0):rows.length;
  $('statReady').textContent=ready.length;$('statError').textContent=job?.counts?.BLOCKED??errors.length;
  $('statSent').textContent=job?.counts?.ACCEPTED??rows.filter(r=>r.status==='ACCEPTED').length;
  $('readySummary').textContent=`${ready.length} klar · ${errors.length} med fejl`;
  $('selectionLabel').textContent=`${selected.length} modtagere valgt`;
  $('sendSelected').disabled=!editable||!selected.length;
  $('cancelJob').classList.toggle('hidden',!job||job.cleaned);$('cancelJob').disabled=!!job&&(job.state==='ANALYZING'||job.state==='SENDING'||job.test_status==='SENDING');
  $('systemStatus').textContent=job?(job.progress||labels[job.state]):'Klar til import';
  const progress=[job?.progress,job?.error,job?.test_status?`Testmail: ${statusLabel(job.test_status)||job.test_status}${job.test_error?' · '+job.test_error:''}`:''].filter(Boolean).join(' · ');
  $('progress').textContent=progress;$('progress').classList.toggle('hidden',!progress);
  const warnings=job?.warnings||[];
  $('warningPanel').classList.toggle('hidden',!warnings.length);$('warningSummary').textContent=`${warnings.length} bemærkninger til Excel-filen`;
  $('warningList').replaceChildren(...warnings.map(w=>text('p',`${w.sheet||'Projektmappe'} · ${w.cell}: ${w.message}`)));
  $('tableEmpty').classList.toggle('hidden',rows.length>0);
  if(job?.cleaned){$('tableEmpty').replaceChildren(text('h3',labels[job.state]),text('p',`${job.counts?.ACCEPTED||0} mails ${job.mail_method==='outlook'?'afleveret til Outlook':'accepteret af Microsoft Graph'}. Personoplysninger og arbejdsfiler er slettet.`));}
  else if(!rows.length){$('tableEmpty').replaceChildren(text('h3',job?.state==='ANALYZING'?'Excel-filen behandles…':'Start med månedens Excel-fil'),text('p','Her får du overblik over modtagere, PDF’er og eventuelle fejl.'));}
  const tbody=$('employees');tbody.replaceChildren();
  for(const r of rows){
    const tr=document.createElement('tr');tr.className='employee-row '+(r.errors.length?'blocked':'ready')+(r.id===activeId?' selected-row':'');
    const checkCell=document.createElement('td');checkCell.className='checkbox-col';const check=document.createElement('input');check.type='checkbox';check.checked=r.selected&&r.status==='READY';check.disabled=!editable||!sendable(r);check.setAttribute('aria-label','Vælg '+r.name);
    check.addEventListener('change',()=>action(async()=>{await saveDraft();job=await api(`/jobs/${job.id}/recipients/${r.id}`,'PATCH',{selected:check.checked});render(false);}));checkCell.append(check);tr.append(checkCell);
    const nameCell=document.createElement('td');nameCell.append(text('span',r.name,'employee-name'),text('span',r.sheet+(r.overview?' ↔ '+r.overview:''),'employee-sheet'));tr.append(nameCell,text('td',r.email||'Ingen adresse'));
    const statusCell=document.createElement('td');const status=r.needs_pdf_review?'Gennemgå PDF':r.needs_review?'Gennemgå besked':statusLabel(r.status);statusCell.append(text('span',status,'status-pill '+(r.status==='ACCEPTED'?'status-sent':r.errors.length?'status-error':r.needs_review||r.needs_pdf_review||['UNKNOWN','FAILED'].includes(r.status)?'status-warning':'status-ready')));tr.append(statusCell);
    tr.addEventListener('click',e=>{if(e.target.tagName==='INPUT')return;action(async()=>{await saveDraft();activeId=r.id;render();});});tbody.append(tr);
  }
  $('selectAll').disabled=!editable||!ready.length;$('selectAll').checked=!!ready.length&&ready.every(r=>r.selected);
  $('errorPanel').classList.toggle('hidden',!errors.length);$('errorSummary').textContent=`${errors.length} medarbejdere / ark kræver rettelse`;
  $('errorList').replaceChildren();
  for(const r of errors){const item=text('div','','error-item');const link=text('button',r.name,'text-btn');link.addEventListener('click',()=>action(async()=>{await saveDraft();activeId=r.id;render();$('detail').scrollIntoView({block:'nearest'});}));item.append(link);for(const error of r.errors){item.append(text('p',error.message));if(error.cells.length)for(const c of error.cells)item.append(text('div',`${c.sheet}!${c.cell}: ${c.value||'(tom)'}`,'cell-reference'));else item.append(text('div','Ark: '+error.sheets.join(', '),'cell-reference'));}$('errorList').append(item);}
  if(refreshDetail&&!dirty&&!saving)renderDetail();
}
function renderDetail(){
  const r=row();$('detailEmpty').classList.toggle('hidden',!!r);$('detail').classList.toggle('hidden',!r);if(!r)return;
  $('detailName').textContent=r.name;$('detailEmail').textContent=r.email||'Ingen gyldig mailadresse';$('detailSheet').textContent=`Ark: ${r.sheet}${r.overview?' · Oversigt: '+r.overview:''}`;
  $('detailStatus').textContent=statusLabel(r.status);$('detailStatus').className='status-pill '+(r.errors.length?'status-error':r.status==='ACCEPTED'?'status-sent':'status-ready');
  $('rowWarnings').replaceChildren(...r.warnings.map(w=>text('p',w)));
  const path=`/api/jobs/${job.id}/recipients/${r.id}/pdf`;$('pdfOpen').href=path;$('pdfDownload').href=path+'?download=true';$('pdfOpen').classList.toggle('hidden',!r.sendable);$('pdfDownload').classList.toggle('hidden',!r.sendable);
  $('subject').value=r.subject;$('body').value=r.body;const editable=canEdit()&&!locked(r);
  ['subject','body','saveMessage','resetMessage'].forEach(id=>$(id).disabled=!editable);
  $('reviewBox').classList.toggle('hidden',!r.needs_review&&!r.needs_pdf_review);$('markReviewed').disabled=!editable;
  $('reviewText').textContent=r.needs_pdf_review?'Excel-filen har advarsler. Åbn PDF’en og kontrollér de nævnte celler, udseendet og beskeden.':'Perioden er ændret. Kontrollér din individuelle besked.';
  $('sendOne').disabled=!canEdit()||!sendable(r);$('testMail').disabled=!canEdit()||!sendable(r);
  $('reconcileBox').classList.toggle('hidden',!['UNKNOWN','FAILED'].includes(r.status));
  $('saveStatus').textContent=r.custom?'Personlig besked':'Fra skabelon';
}
async function openJob(id){await saveDraft();job=await api('/jobs/'+id);activeId=null;chosenFile=null;dirty=false;ccDirty=false;setPeriod(job.month,job.year);await tab('overview');render();}
async function history(){const jobs=await api('/jobs');$('historyList').replaceChildren();if(!jobs.length)$('historyList').append(text('p','Der er endnu ingen kørsler.'));for(const j of jobs){const card=text('article','','history-card');const detail=document.createElement('div');detail.append(text('h3',`${j.kind==='fastansat'?'Fastansatte':'Vikarer'} · ${months[j.month-1]} ${j.year}`),text('p',`${labels[j.state]} · ${methodName(j.mail_method)} · Oprettet ${displayDate(j.created)}${j.filename?' · '+j.filename:''}`));const button=text('button',j.cleaned?'Vis kvittering':'Åbn kørsel','secondary-btn');button.onclick=()=>action(()=>openJob(j.id));card.append(detail,button);$('historyList').append(card);}}
function fillSettings(){const s=settings;$('mailMethod').value=s.mail_method||'graph';outlookAccounts=s.outlook_account?[s.outlook_account]:[];showOutlookAccounts(s.outlook_account?.id);showMethod();$('clientId').value=s.client_id;$('tenantId').value=s.tenant_id;$('loProgram').value=s.libreoffice_program;$('loPython').value=s.libreoffice_python;$('templateFastSubject').value=s.templates.fastansat.subject;$('templateFastBody').value=s.templates.fastansat.body;$('templateVikarSubject').value=s.templates.vikar.subject;$('templateVikarBody').value=s.templates.vikar.body;}
async function saveSettings(){const updatedSettings={mail_method:$('mailMethod').value,outlook_account:outlookAccounts.find(a=>a.id===$('outlookAccount').value)||null,client_id:$('clientId').value.trim(),tenant_id:$('tenantId').value.trim()||'organizations',libreoffice_program:$('loProgram').value.trim(),libreoffice_python:$('loPython').value.trim(),templates:{fastansat:{subject:$('templateFastSubject').value,body:$('templateFastBody').value},vikar:{subject:$('templateVikarSubject').value,body:$('templateVikarBody').value}}};await api('/settings','PUT',updatedSettings);settings=updatedSettings;account=(await api('/auth/status')).account;showAccount();}
let approvalIds=[],approvalRevision=0;
async function previewSend(ids){await saveDraft();job=await api('/jobs/'+job.id);account=(await api('/auth/status')).account;showAccount();if(!account)throw new Error('Vælg og gem afsenderkonto under Indstillinger først.');if((job.mail_method||'graph')!==settings.mail_method)throw new Error('Denne kørsel bruger '+methodName(job.mail_method)+'. Vælg samme metode under Indstillinger.');approvalIds=ids;approvalRevision=job.revision;$('approvalSender').textContent='Afsender: '+account.email+' · '+methodName(job.mail_method);$('approvalPeriod').textContent='Lønperiode: '+job.period;$('approvalCc').textContent=job.cc?'Cc på hver mail: '+job.cc:'Cc: Ingen';$('approvalCount').textContent=ids.length+' modtagere';$('approvalRecipients').replaceChildren();$('approvalExcluded').replaceChildren();for(const r of job.recipients){const included=ids.includes(r.id);$(included?'approvalRecipients':'approvalExcluded').append(text('li',`${r.name} · ${r.email||'Ingen adresse'}${included?'':' · '+(r.needs_pdf_review?'PDF kræver gennemgang':r.needs_review?'Besked kræver gennemgang':r.status==='READY'?'Ikke valgt':statusLabel(r.status))}`));}if(ids.length===job.recipients.length)$('approvalExcluded').append(text('li','Ingen'));$('approvalDialog').showModal();}
function wire(){
  $('mailMethod').onchange=showMethod;
  $('loadOutlook').onclick=()=>action(async()=>{$('outlookStatus').textContent='Kontakter klassisk Outlook…';$('loadOutlook').disabled=true;try{const previous=$('outlookAccount').value;const result=await api('/outlook/accounts','POST');outlookAccounts=result.accounts;showOutlookAccounts(previous);$('outlookStatus').textContent=outlookAccounts.length?'Vælg afsenderkontoen og tryk Gem indstillinger.':'Ingen egnede konti fundet. Tilføj afsenderkontoen i klassisk Outlook.';}catch(e){$('outlookStatus').textContent=e.message;throw e;}finally{$('loadOutlook').disabled=false;}});

  document.querySelectorAll('[data-tab]').forEach(el=>el.onclick=()=>action(()=>tab(el.dataset.tab)));$('accountSettings').onclick=()=>action(()=>tab('settings'));
  $('chooseFile').onclick=()=>$('fileInput').click();$('fileInput').onchange=()=>{chosenFile=$('fileInput').files[0];render();};
  $('analyze').onclick=()=>action(async()=>{if(!chosenFile)throw new Error('Vælg først en .xlsx-fil.');const form=new FormData();form.append('file',chosenFile);form.append('kind',$('kind').value);form.append('month',$('month').value);form.append('year',$('year').value);form.append('cc',$('cc').value);$('analyze').disabled=true;try{job=await api('/jobs','POST',form);ccDirty=false;render();}finally{if(!job)$('analyze').disabled=false;}});
  $('cc').oninput=()=>{ccDirty=true;$('ccStatus').textContent=job?'Ikke gemt':'Gemmes ved import';};
  $('ccForm').onsubmit=e=>{e.preventDefault();action(async()=>{await saveDraft();render(false);});};
  $('newJob').onclick=()=>action(async()=>{await saveDraft();job=null;activeId=null;chosenFile=null;dirty=false;ccDirty=false;$('fileInput').value='';render();});
  $('shutdown').onclick=()=>action(async()=>{await saveDraft();await api('/shutdown','POST');stopped=true;notice('Appen er afsluttet. Du kan lukke denne browserfane. Kladderne er gemt.',true);$('systemStatus').textContent='Appen er afsluttet';});
  for(const id of ['month','year'])$(id).onchange=()=>action(async()=>{await saveDraft();try{if(job){job=await api(`/jobs/${job.id}/period`,'PATCH',{month:Number($('month').value),year:Number($('year').value)});render();}$('periodLabel').textContent=currentPeriod();}catch(e){if(job)setPeriod(job.month,job.year);throw e;}});
  for(const id of ['subject','body'])$(id).oninput=()=>{dirty=true;$('saveStatus').textContent='Ikke gemt';clearTimeout(timer);timer=setTimeout(()=>saveDraft().catch(e=>notice(e.message)),900);};
  $('messageForm').onsubmit=e=>{e.preventDefault();action(saveDraft);};
  $('resetMessage').onclick=()=>action(async()=>{await saveDraft();job=await api(`/jobs/${job.id}/recipients/${activeId}`,'PATCH',{reset:true});render();});
  $('markReviewed').onclick=()=>action(async()=>{await saveDraft();job=await api(`/jobs/${job.id}/recipients/${activeId}`,'PATCH',{reviewed:true});render();});
  $('selectAll').onchange=()=>action(async()=>{await saveDraft();const selected=$('selectAll').checked;for(const r of job.recipients.filter(sendable))job=await api(`/jobs/${job.id}/recipients/${r.id}`,'PATCH',{selected});render(false);});
  $('sendSelected').onclick=()=>action(()=>previewSend(job.recipients.filter(r=>sendable(r)&&r.selected).map(r=>r.id)));
  $('sendOne').onclick=()=>action(()=>previewSend([activeId]));$('closeApproval').onclick=()=>$('approvalDialog').close();
  $('confirmSend').onclick=()=>action(async()=>{$('confirmSend').disabled=true;try{const approval=await api(`/jobs/${job.id}/approve`,'POST',{ids:approvalIds,revision:approvalRevision});job=await api(`/jobs/${job.id}/send`,'POST',{approval_id:approval.id});$('approvalDialog').close();render();}catch(e){$('approvalDialog').close();throw e;}finally{$('confirmSend').disabled=false;}});
  $('testMail').onclick=()=>action(async()=>{await saveDraft();account=(await api('/auth/status')).account;if(!account)throw new Error('Vælg og gem afsenderkonto under Indstillinger først.');if(!confirm(`Send en test med denne persons PDF til dig selv (${account.email})?`))return;job=await api(`/jobs/${job.id}/recipients/${activeId}/test`,'POST');render();});
  for(const [id,outcome] of [['markAccepted','accepted'],['markRetry','retry']])$(id).onclick=()=>action(async()=>{if(!confirm(outcome==='retry'?'Har du kontrolleret både Udbakke og Sendt post, og sikret at mailen ikke allerede er afsendt eller venter i Udbakke?':'Bekræft, at du har fundet mailen i Sendt post.'))return;job=await api(`/jobs/${job.id}/recipients/${activeId}/reconcile`,'POST',{outcome});render();});
  $('cancelJob').onclick=()=>action(async()=>{await saveDraft();if(!confirm('Afslut kørslen og slet dens Excel-fil, PDF’er og beskeder? Allerede afsendte mails kan ikke trækkes tilbage.'))return;job=await api(`/jobs/${job.id}/cancel`,'POST');render();});
  $('settingsForm').onsubmit=e=>{e.preventDefault();action(async()=>{await saveSettings();notice('Indstillinger gemt. Eksisterende kladder beholder deres skabeloner.',true);});};
  $('connect').onclick=()=>action(async()=>{await saveSettings();const login=await api('/auth/connect','POST');$('loginCode').textContent=login.user_code;$('loginLink').href=login.url;$('loginBox').classList.remove('hidden');notice('Åbn Microsoft-login og indtast koden.',true);});
  $('logout').onclick=()=>action(async()=>{await api('/auth/logout','POST');account=null;showAccount();$('loginBox').classList.add('hidden');});
  $('copyCode').onclick=()=>action(async()=>{await navigator.clipboard.writeText($('loginCode').textContent);notice('Login-kode kopieret.',true);});
  $('checkRuntime').onclick=()=>action(async()=>{await saveSettings();$('runtimeStatus').textContent='Kontrollerer…';const r=await api('/runtime');$('runtimeStatus').textContent=`Klassisk Outlook: ${r.outlook_classic?'fundet':'ikke fundet'} · Microsoft Excel: ${r.excel?'fundet':'ikke fundet'} · LibreOffice/UNO: ${r.libreoffice?'klar':r.libreoffice_error}`;});
  window.addEventListener('beforeunload',e=>{if(dirty||saving||ccDirty||ccSaving){e.preventDefault();e.returnValue='';}});
}
async function poll(){
  if(stopped)return;
  try{
    if(!busy&&!dirty&&!saving&&!ccDirty&&!ccSaving){
      if(job&&!job.cleaned){const updated=await api('/jobs/'+job.id);const changed=JSON.stringify(updated)!==JSON.stringify(job);job=updated;if(changed)render(!['subject','body'].includes(document.activeElement?.id));}
      const auth=await api('/auth/status');account=auth.account;showAccount();
      if(auth.login?.status==='connected')$('loginBox').classList.add('hidden');
      if(auth.login?.status==='error'){$('connectionStatus').textContent=auth.login.error;$('loginBox').classList.add('hidden');}
    }
  }catch(e){$('systemStatus').textContent='Forbindelsen til appen er afbrudt';}finally{setTimeout(poll,2000);}
}
async function init(){try{const session=await api('/session');csrf=session.csrf;account=session.account;months.forEach((m,i)=>$('month').append(new Option(m[0].toUpperCase()+m.slice(1),i+1)));for(let y=session.year-5;y<=session.year+5;y++)$('year').append(new Option(y,y));setPeriod(session.month,session.year);settings=await api('/settings');fillSettings();wire();showAccount();render();poll();}catch(e){notice('Appen kunne ikke startes: '+e.message);}}
init();
