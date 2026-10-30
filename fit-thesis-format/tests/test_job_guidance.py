"""Recovery contract tests; no simulated renderer is native Office proof."""
import copy
import json
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from job_guidance import repair_plan


class RecoveryContractTests(unittest.TestCase):
    def test_long_inspection_and_findings_do_not_fill_stdout_with_full_maps(self):
        import thesis_job as J
        candidates=[{'id':f'p{i:06d}','role':'chapter','text':'完整标题'+str(i)+'甲'*100,'reason':'需确认'}
                    for i in range(1000)]
        state={'status':'semantic_review_required','source_sha256':'hash','delivery_ready':False,
               'inspection':{'source_sha256':'hash','counts':{'paragraphs':1000},'candidates':candidates,
                             'map':{'paragraphs':{c['id']:{'expected_text':'正文'*1000} for c in candidates}}},
               'blocking_items':[{'check':'format','issues':[
                   {'category':'heading.long_text_requires_review','actual':'原始证据'*2000} for _ in range(300)]}]}
        before=copy.deepcopy(state)
        result=J.response(Path('/job'),state)
        self.assertLess(len(json.dumps(result,ensure_ascii=False).encode('utf-8')),60000)
        task=result['next'][0]
        self.assertNotIn('map',task['inspection'])
        self.assertEqual(task['inspection_file'],str(Path('/job')/'inspection.json'))
        self.assertEqual(task['inspection']['candidate_count'],1000)
        self.assertLessEqual(len(task['inspection']['candidates']),30)
        self.assertTrue(task['inspection']['sample_is_not_full_coverage'])
        self.assertEqual(task['inspection']['candidates'][0]['text'],candidates[0]['text'])
        self.assertTrue(result['repair_plan']['completion_policy']['must_continue_format_repairs'])
        self.assertEqual(result['blocking_item_count'],1)
        self.assertEqual(result['blocking_items_omitted'],1)
        self.assertEqual(state,before)

    def test_oversized_candidate_is_left_in_file_without_text_truncation(self):
        import thesis_job as J
        long_text='不截断的原稿文字'*10000
        state={'status':'semantic_review_required','source_sha256':'hash',
               'inspection':{'candidates':[{'id':'p000001','text':long_text,'role':'chapter'},
                                           {'id':'p000002','text':'已经确认','role':'chapter'},
                                           {'id':'p000003','text':'完整短标题','role':'chapter'}]},
               'decisions':{'paragraphs':{'p000002':{'role':'body'}}}}
        task=J.next_tasks(Path('/job'),state)[0]
        self.assertEqual(task['inspection']['pending_candidate_count'],2)
        self.assertEqual(task['inspection']['sample_omitted_count'],1)
        self.assertEqual(task['inspection']['candidates'],[{'id':'p000003','role':'chapter','text':'完整短标题'}])
        self.assertEqual(state['inspection']['candidates'][0]['text'],long_text)
        self.assertTrue(task['inspection']['complete_review_required'])

    def test_public_response_exposes_recovery_and_full_decision_merge(self):
        import thesis_job as J
        state = {'status': 'correction_required', 'delivery_ready': False, 'source_sha256': 'hash',
                 'inspection': {'candidates': []}, 'blocking_items': [{'category': 'toc.depth'}]}
        response = J.response(Path('/job'), state)
        self.assertTrue(response['ok'])
        self.assertFalse(response['delivery_ready'])
        self.assertTrue(response['repair_plan']['completion_policy']['must_continue_format_repairs'])
        task = response['next'][0]
        self.assertTrue(task['existing_decisions']['merge_before_submit'])
        self.assertTrue(task['decision_template_is_example_only'])
        self.assertEqual(task['submit_request']['kind'], 'decisions')

    def test_manifest_entry_dispatches_to_same_state_machine(self):
        import agent_cli as A
        import thesis_job as J
        manifest = json.loads((Path(__file__).resolve().parents[1] / 'agent-tools.json').read_text(encoding='utf-8'))
        self.assertEqual(manifest['transport']['command'][1], 'scripts/agent_cli.py')
        for action in manifest['execution_contract']['workflow_actions']:
            request = {'action': action, 'work_dir': '/job'}
            with patch.object(J, 'execute', return_value={'sentinel': action}) as call:
                self.assertEqual(A.execute(request), {'sentinel': action})
                call.assert_called_once_with(request)

    def test_no_progress_next_is_actionable(self):
        import thesis_job as J
        tasks = J.next_tasks(Path('/job'), {'status': 'no_progress', 'blocking_items': [{'reason': 'same'}]})
        self.assertEqual(tasks[0]['kind'], 'diagnose_repeated_failure')
        self.assertEqual(tasks[0]['issues'][0]['reason'], 'same')

    def test_content_gaps_do_not_hide_format_repairs(self):
        state = {'status': 'correction_required', 'delivery_ready': False,
                 'source_sha256': 'hash', 'content_advice': [{'code': 'missing_identity_fields'}],
                 'verification': {'blocking_items': [{'check': 'format', 'issues': [
                     {'category': 'toc.cached_number_prefix_mismatch'},
                     {'category': 'front.en_title_missing'}]}]}}
        before = copy.deepcopy(state)
        plan = repair_plan(state, '/job')
        self.assertEqual(plan['action'], 'repair_and_reverify')
        self.assertEqual(plan['counts']['missing_content'], 2)
        self.assertEqual(plan['counts']['format_repairs'], 1)
        self.assertTrue(plan['completion_policy']['must_continue_format_repairs'])
        self.assertFalse(plan['completion_policy']['may_claim_complete'])
        self.assertEqual(state, before)

    def test_years_and_unrecognized_sections_require_local_evidence_first(self):
        state = {'status': 'correction_required', 'unresolved': [
            {'category': 'citation.missing_year'},
            {'category': 'structure.acknowledgement_missing_or_unconfirmed'}]}
        plan = repair_plan(state, '/job')
        self.assertEqual(plan['counts']['evidence_checks'], 2)
        self.assertEqual(plan['counts']['missing_content'], 0)

    def test_unknown_issue_is_not_silently_dropped(self):
        plan = repair_plan({'status': 'correction_required', 'blocking_items': [
            {'code': 'future.unrecognized'}]}, '/job')
        self.assertEqual(plan['counts']['evidence_checks'], 1)
        self.assertTrue(plan['completion_policy']['must_continue_format_repairs'])

    def test_no_progress_routes_diagnosis_and_preserves_blocker(self):
        plan = repair_plan({'status': 'no_progress', 'blocking_items': [
            {'reason': 'repeated identical submission'}]}, '/job')
        self.assertEqual(plan['action'], 'diagnose_repeated_failure')
        self.assertFalse(plan['completion_policy']['may_claim_complete'])
        self.assertEqual(plan['counts']['evidence_checks'], 1)

    def test_only_actual_complete_state_permits_completion(self):
        for state in ({'status': 'correction_required', 'delivery_ready': True},
                      {'status': 'complete', 'delivery_ready': False}):
            self.assertFalse(repair_plan(state, '/job')['completion_policy']['may_claim_complete'])
        plan = repair_plan({'status': 'complete', 'delivery_ready': True}, '/job')
        self.assertTrue(plan['completion_policy']['may_claim_complete'])

    def test_missing_capability_does_not_request_new_run(self):
        plan = repair_plan({'status': 'capability_required', 'blocking_items': [
            {'check': 'dual_engine'}]}, '/job with spaces')
        self.assertEqual(plan['counts']['capability_repairs'], 1)
        self.assertEqual(plan['commands']['resume'],
                         {'action': 'resume', 'work_dir': '/job with spaces'})
        self.assertNotIn('run', plan['commands'])

    def test_duplicate_findings_do_not_inflate_repair_queue(self):
        finding = {'category': 'paragraph.format', 'location': {'paragraph_id': 'p000123'}}
        plan = repair_plan({'status': 'correction_required',
                            'blocking_items': [finding], 'unresolved': [finding]}, '/job')
        self.assertEqual(plan['counts']['format_repairs'], 1)


if __name__ == '__main__':
    unittest.main()
