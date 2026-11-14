/**
 * Clash of Classes -- Apps Script backend
 * File 35: site content that used to be built into the page, now on Sheet tabs (edit them in
 * Admin › Sheet or in the Sheet itself).
 *
 * LRG TRAITS tab: Order | Month | Trait | Statement | Skill | Skill Text. One row per skill; the
 *   first row of a month carries the Trait name and its "I ..." statement (blank below = same month).
 *   The Trait page, the home page trait, the TV and the trait tracker all read it.
 * WHEEL SLICES tab: Order | Color | Kind | Value | Name | Loyalty. One row per slice of the Spin
 *   Wheel, in order around the wheel. Kind pts = that many points; special = a named slice (NONE,
 *   LOYALTY...). Loyalty ticked = it only pays the squad whose color it is.
 * Run setup fills both with what the site showed before. A tab left empty = the built-in list.
 */

const TRAIT_HEADERS_ = ['Order', 'Month', 'Trait', 'Statement', 'Skill', 'Skill Text'];
const WHEEL_HEADERS_ = ['Order', 'Color', 'Kind', 'Value', 'Name', 'Loyalty'];
const TRAIT_SEED_ = [
  [1, 'September', 'Continual Learning & Growth Mindset', '"I believe that intelligence and ability are not fixed traits and can develop over time."', 'Awareness', 'Acknowledge that new learning is achieved and maintained through effort and continual practice.'],
  [2, 'September', '', '', 'Mindset', 'Open to and embrace challenges.'],
  [3, 'September', '', '', 'Feedback', 'Seek and utilize feedback throughout the learning process to continue to improve.'],
  [4, 'September', '', '', 'Persistence', 'Give sustained effort in learning over time when encountering a challenging concept or topic.'],
  [5, 'October', 'Critical Thinking & Problem Solving', '"I think through complex issues to identify and make informed decisions."', 'Gathering Information', 'Collect relevant information, data, and perspectives to generate a deeper understanding of the problem.'],
  [6, 'October', '', '', 'Analyzing Information', 'Engage in a deliberate process to determine information\'s accuracy and relevance in order to draw conclusions.'],
  [7, 'October', '', '', 'Identifying Solutions', 'Consider a variety of possibilities to make sound judgments and decisions about action.'],
  [8, 'October', '', '', 'Taking Action', 'Make an informed decision and use available resources to take action on a possible solution.'],
  [9, 'October', '', '', 'Reflecting', 'Interpret the outcome of the action and consider possible adjustments for improvement.'],
  [10, 'November', 'Creativity & Innovation', '"I engage in an ongoing, cyclical process of design and improvement."', 'Original Ideas', 'Generate original ideas bringing unique perspectives and imagining new opportunities.'],
  [11, 'November', '', '', 'Techniques', 'Utilize a variety of techniques to imagine, visualize, brainstorm, and get feedback.'],
  [12, 'November', '', '', 'Iteration', 'Use feedback to evaluate the implementation of an idea; effective design is ongoing and adaptive.'],
  [13, 'November', '', '', 'Implementation', 'Understand situational limits when moving from idea to action.'],
  [14, 'December', 'Honesty, Integrity & Responsibility', '"I am honest and take ownership of my actions, regardless of the circumstances."', 'Consistency', 'Align actions with words; follow through.'],
  [15, 'December', '', '', 'Responsibility', 'Own one\'s intentions, actions, and the results of those actions.'],
  [16, 'December', '', '', 'Principles', 'Act morally and ethically, even when no one is looking.'],
  [17, 'January', 'Collaboration & Teamwork', '"I effectively collaborate with others, and compromise when necessary for the betterment of the group."', 'Collaboration', 'Engage in interdependent planning and decision-making to accomplish a common goal.'],
  [18, 'January', '', '', 'Dynamics', 'Identify and appropriately use the skill set of each individual to support the team.'],
  [19, 'January', '', '', 'Compromise', 'Value and incorporate the ideas, perspectives, strengths, and expertise of others.'],
  [20, 'January', '', '', 'Accountability', 'Demonstrate trust in team members and assume mutual accountability for agreed-upon responsibilities.'],
  [21, 'February', 'Communication & Empathy', '"I effectively communicate in oral, written, visual, and nonverbal forms, with a variety of audiences."', 'Active Listening', 'Acknowledge others\' feelings and opinions through engagement and thoughtful processing of information.'],
  [22, 'February', '', '', 'Meaningful Inquiry', 'Ask clarifying questions to develop a broader understanding of the concept, problem, or situation.'],
  [23, 'February', '', '', 'Audience', 'Demonstrate an awareness of audience members\' identity and background knowledge to inform communication strategies.'],
  [24, 'February', '', '', 'Clear Purpose', 'Express ideas with the purpose of informing, instructing, motivating, or persuading using appropriate language and format.'],
  [25, 'March', 'Healthy Living, Wellness & Self-Awareness', '"I make healthy choices and develop an understanding of and confidence in my own identity."', 'Physical Health & Safety', 'Develop a healthy lifestyle through diet, nutrition, exercise, sleep, and safe choices.'],
  [26, 'March', '', '', 'Self-Awareness', 'Develop an understanding of and confidence in my characteristics, interests, strengths, and areas for growth.'],
  [27, 'March', '', '', 'Goal Setting & Management', 'Plan, organize, prioritize, and initiate progress on specific tasks in pursuit of a goal.'],
  [28, 'March', '', '', 'Emotional Regulation', 'Use effective strategies to express emotions and manage responses appropriately to the situation.'],
  [29, 'March', '', '', 'Social Awareness', 'Recognize, understand, and respectfully acknowledge similarities and differences in others\' perspectives.'],
  [30, 'March', '', '', 'Relationships', 'Establish, maintain, and repair healthy relationships.'],
  [31, 'March', '', '', 'Responsible Decision Making', 'Consider various options and anticipate possible outcomes to make caring and constructive choices.'],
  [32, 'March', '', '', 'Advocacy', 'Recognize when help is needed and seek appropriate people and resources to support myself or others.'],
  [33, 'April', 'Resilience & Grit', '"I persist through adversity and never give up on myself or the team."', 'Dedication', 'Use courage and determination to consistently work toward goals.'],
  [34, 'April', '', '', 'Perseverance', 'Respond positively to unexpected challenges and rebound from setbacks.'],
  [35, 'April', '', '', 'Reflection', 'Consider experiences and use feedback from challenges as opportunities for improvement.'],
  [36, 'May', 'Adaptability & Flexibility', '"I adjust to new or different circumstances with a solution-focused attitude."', 'Perspectives', 'Seek to understand the situation and the options and strategies available.'],
  [37, 'May', '', '', 'Adjustments', 'Adjust to new or different circumstances by changing thinking, behaviors, actions, contexts, and/or objectives.'],
  [38, 'May', '', '', 'Frame of Mind', 'Maintain a solution-focused attitude when conditions change.'],
];
const WHEEL_SEED_ = [
  [1, 'BLACK', 'pts', 1, '', false],
  [2, 'BLACK', 'pts', 10, '', false],
  [3, 'BLACK', 'pts', 25, '', false],
  [4, 'BLACK', 'special', 0, 'NONE', false],
  [5, 'BLACK', 'pts', 50, '', false],
  [6, 'BLACK', 'pts', 5, '', false],
  [7, 'GOLD', 'pts', 20, '', false],
  [8, 'GOLD', 'pts', 5, '', false],
  [9, 'GOLD', 'special', 20, 'LOYALTY', true],
  [10, 'GOLD', 'pts', 50, '', false],
  [11, 'GOLD', 'pts', 10, '', false],
  [12, 'GOLD', 'pts', 25, '', false],
  [13, 'GREY', 'pts', 1, '', false],
  [14, 'GREY', 'pts', 30, '', false],
  [15, 'GREY', 'pts', 10, '', false],
  [16, 'GREY', 'special', 0, 'ZERO', false],
  [17, 'GREY', 'pts', 20, '', false],
  [18, 'GREY', 'pts', 75, '', false],
  [19, 'PURPLE', 'pts', 5, '', false],
  [20, 'PURPLE', 'pts', 50, '', false],
  [21, 'PURPLE', 'special', 0, 'AUTO', false],
  [22, 'PURPLE', 'pts', 25, '', false],
  [23, 'PURPLE', 'pts', 10, '', false],
  [24, 'PURPLE', 'special', 20, 'LOYALTY', true],
];

function ensureSiteContentTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['LRG Traits', TRAIT_HEADERS_, TRAIT_SEED_], ['Wheel Slices', WHEEL_HEADERS_, WHEEL_SEED_]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]);
  });
}

/** [{ month, trait, desc, skills: [{ label, desc }] }] from the LRG Traits tab, or null when it is empty. */
function siteTraits_() {
  const rows = sheetRows_('LRG Traits').filter(function (r) { return String(r['Month'] || '').trim(); })
    .sort(function (a, b) { return (Number(a['Order']) || 0) - (Number(b['Order']) || 0); });
  if (!rows.length) return null;
  const out = [];
  rows.forEach(function (r) {
    const month = String(r['Month']).trim();
    let t = out.filter(function (x) { return x.month === month; })[0];
    if (!t) { t = { month: month, trait: '', desc: '', skills: [] }; out.push(t); }
    if (String(r['Trait'] || '').trim()) t.trait = String(r['Trait']).trim();
    if (String(r['Statement'] || '').trim()) t.desc = String(r['Statement']).trim();
    if (String(r['Skill'] || '').trim()) t.skills.push({ label: String(r['Skill']).trim(), desc: String(r['Skill Text'] || '').trim() });
  });
  const keep = out.filter(function (t) { return t.trait; });
  return keep.length ? keep : null;
}

/** [{ color, type, val, name, loyalty }] from the Wheel Slices tab, or null when it is empty. */
function siteWheel_() {
  const rows = sheetRows_('Wheel Slices').filter(function (r) { return SQUAD_KEYS_.indexOf(squadOf_(r['Color'])) > -1; })
    .sort(function (a, b) { return (Number(a['Order']) || 0) - (Number(b['Order']) || 0); });
  if (rows.length < 4) return null;
  return rows.map(function (r) {
    const special = /special/i.test(String(r['Kind'] || ''));
    const o = { color: squadOf_(r['Color']), type: special ? 'special' : 'pts', val: Number(r['Value']) || 0 };
    if (String(r['Name'] || '').trim()) o.name = String(r['Name']).trim().toUpperCase();
    if (ticked_(r['Loyalty'])) o.loyalty = true;
    return o;
  });
}
