import React, { useState } from 'react';
import './EnhancementRoadmap.css';

// ── ROADMAP DATA MODEL ────────────────────────────────────────────────────────
const ROADMAP_ITEMS = [
  {
    id: 'trend-lines',
    titleEn: 'Trend Lines over Time',
    titleAr: 'مخططات الاتجاهات مع الوقت',
    quadrant: 'quick-wins',
    phase: 1,
    impact: 'High',
    effort: 'Low',
    status: 'Proposed',
    dependencies: [],
    descEn: 'Add historical performance charts showing daily/weekly trends instead of single static snapshots.',
    descAr: 'إضافة مخططات الأداء التاريخي لتوضيح الاتجاهات اليومية والأسبوعية بدلاً من اللقطات الثابتة فقط.',
    businessEn: 'Helps sales directors spot early drops in team productivity and see long-term behavioral changes.',
    businessAr: 'تساعد مدراء المبيعات في رصد الانخفاض المبكر في الإنتاجية وتحديد التغيرات السلوكية طويلة المدى.',
    stepsEn: ['Integrate Recharts components', 'Fetch historical daily metrics from supabase', 'Design multi-line trend overlay.'],
    stepsAr: ['دمج مكونات Recharts', 'جلب المقاييس اليومية التاريخية من Supabase', 'تصميم تراكب اتجاه متعدد الخطوط.']
  },
  {
    id: 'benchmarking',
    titleEn: 'Comparative Benchmarking',
    titleAr: 'المقارنة المعيارية للأداء',
    quadrant: 'quick-wins',
    phase: 1,
    impact: 'High',
    effort: 'Low',
    status: 'Proposed',
    dependencies: ['progress-cards'],
    descEn: 'Compare each sales rep against team/company averages, using color-coded badges (Exceeding / Meeting / Below average).',
    descAr: 'مقارنة كل مندوب مبيعات بمتوسطات الفريق/الشركة باستخدام شارات ملونة (يتجاوز / يلبي / أقل من المتوسط).',
    businessEn: 'Promotes healthy team competition and instantly highlights who needs immediate operational assistance.',
    businessAr: 'يعزز المنافسة الصحية بين الفريق ويحدد فوراً من يحتاج إلى مساعدة ميدانية عاجلة.',
    stepsEn: ['Compute team averages dynamically', 'Apply conditional color mapping (Green/Yellow/Red)', 'Render ranking badge on KPI view.'],
    stepsAr: ['حساب متوسطات الفريق ديناميكياً', 'تطبيق تصنيف الألوان المشروط (أخضر/أصفر/أحمر)', 'عرض شارة الترتيب في لوحة المؤشرات.']
  },
  {
    id: 'progress-cards',
    titleEn: 'Target Progress Bars',
    titleAr: 'أشرطة التقدم نحو الأهداف',
    quadrant: 'quick-wins',
    phase: 1,
    impact: 'High',
    effort: 'Low',
    status: 'Proposed',
    dependencies: [],
    descEn: 'KPI cards with progress bar indicators showing actual performance vs monthly and quarterly visit targets.',
    descAr: 'بطاقات مؤشرات الأداء مع أشرطة تقدم توضح الزيارات الفعلية مقارنة بالأهداف الشهرية والربع سنوية.',
    businessEn: 'Aligns field operations directly with strategic targets, visualising target achievement ratios.',
    businessAr: 'يربط العمليات الميدانية مباشرة بالأهداف الاستراتيجية، مما يوضح نسب تحقيق المستهدفات.',
    stepsEn: ['Add target goals metadata', 'Calculate percent progress values', 'Implement CSS progress slider animation.'],
    stepsAr: ['إضافة مستهدفات العمليات كبيانات تعريفية', 'حساب نسب التقدم المئوية', 'تنفيذ حركة شريط التقدم بـ CSS.']
  },
  {
    id: 'quick-filters',
    titleEn: 'Performance Threshold Filters',
    titleAr: 'فلاتر عتبة الأداء السريعة',
    quadrant: 'quick-wins',
    phase: 1,
    impact: 'High',
    effort: 'Low',
    status: 'Proposed',
    dependencies: [],
    descEn: 'Quick selectors for custom date ranges and filters to isolate top or bottom performing reps based on custom thresholds.',
    descAr: 'محددات سريعة لنطاقات التواريخ المخصصة وفلاتر لعزل المندوبين ذوي الأداء الأعلى أو الأقل بناءً على عتبات مخصصة.',
    businessEn: 'Saves executive management valuable analytical time by highlighting outliers in under three clicks.',
    businessAr: 'يوفر على الإدارة التنفيذية وقتاً ثميناً للتحليل من خلال تسليط الضوء على الحالات الشاذة في أقل من 3 نقرات.',
    stepsEn: ['Introduce custom date range input', 'Build slider-based numeric filters', 'Sync selected thresholds to card views.'],
    stepsAr: ['إدخال حقول اختيار تاريخ مخصصة', 'بناء فلاتر رقمية منزلقة للأداء', 'ربط العتبات المحددة مباشرة مع بطاقات العرض.']
  },
  {
    id: 'export-templates',
    titleEn: 'Pre-built Export Templates',
    titleAr: 'قوالب تصدير البيانات الجاهزة',
    quadrant: 'quick-wins',
    phase: 1,
    impact: 'Medium',
    effort: 'Low',
    status: 'Proposed',
    dependencies: [],
    descEn: 'Pre-designed XLSX templates tailored for different stakeholders (Directors, Supervisors, or Rep summaries).',
    descAr: 'قوالب إكسيل جاهزة ومصممة خصيصاً لمختلف الأطراف المعنية (مدراء، مشرفون، أو ملخصات للمندوبين).',
    businessEn: 'Eliminates manual reformatting of report spreadsheets, ensuring unified corporate reporting standards.',
    businessAr: 'يلغي الحاجة لإعادة تنسيق التقارير يدوياً، ويضمن توحيد معايير التقارير المؤسسية.',
    stepsEn: ['Create structured XLSX sheet layouts', 'Implement layout picker option', 'Build Excel formatting triggers with xlsx.'],
    stepsAr: ['إنشاء تخطيطات مخصصة لملفات XLSX', 'تنفيذ خيار اختيار القالب', 'تطبيق منسق الخلايا والمظهر التلقائي.']
  },
  {
    id: 'territory-map',
    titleEn: 'Territory Coverage Map',
    titleAr: 'خريطة تغطية المناطق الجغرافية',
    quadrant: 'strategic',
    phase: 2,
    impact: 'High',
    effort: 'High',
    status: 'Proposed',
    dependencies: ['validation-warnings'],
    descEn: 'Interactive geographic map visualization of clinics, hospitals, and pharmacies, showcasing visited vs unvisited zones.',
    descAr: 'عرض مرئي تفاعلي للمناطق الجغرافية للعيادات، المستشفيات والصيدليات يوضح المناطق التي تمت زيارتها والتي لم تزر.',
    businessEn: 'Reveals literal gaps in physical coverage, maximizing market reach and reducing travel inefficiency.',
    businessAr: 'يكشف الثغرات الجغرافية الفعلية للتغطية، مما يزيد من انتشار السوق ويقلل من عدم كفاءة التنقل.',
    stepsEn: ['Acquire clinic/doctor latitude & longitude data', 'Integrate Leaflet or Mapbox maps', 'Plot colored pins based on visit status.'],
    stepsAr: ['الحصول على إحداثيات العيادات والأطباء (خطوط الطول والعرض)', 'دمج خرائط Leaflet أو Mapbox', 'رسم دبابيس ملونة تعتمد على حالة الزيارة.']
  },
  {
    id: 'call-quality',
    titleEn: 'Call Quality & Outcome Tracking',
    titleAr: 'تتبع جودة ومخرجات الزيارات',
    quadrant: 'strategic',
    phase: 2,
    impact: 'High',
    effort: 'Medium',
    status: 'Proposed',
    dependencies: [],
    descEn: 'Log call outcomes (conversion rates, product samples delivered, doctor feedback and next step commitments).',
    descAr: 'تسجيل مخرجات الزيارة (معدلات التحويل، العينات الطبية الموزعة، مرئيات الأطباء والالتزامات للخطوة القادمة).',
    businessEn: 'Shifts focus from sheer volume of calls to quality and direct sales impact, unlocking actual revenue indicators.',
    businessAr: 'ينقل التركيز من الحجم الكمي للزيارات إلى الجودة وتأثيرها المباشر في المبيعات، كاشفاً مؤشرات الإيرادات الحقيقية.',
    stepsEn: ['Extend database schema for outcome data', 'Update rep reporting forms with sample counts', 'Generate conversion widgets.'],
    stepsAr: ['توسيع هيكل قاعدة البيانات لبيانات المخرجات', 'تحديث نماذج تقارير المندوبين لتشمل كميات العينات', 'إنشاء لوحات لمعدلات التحويل.']
  },
  {
    id: 'coaching-engine',
    titleEn: 'AI Coaching Recommendations',
    titleAr: 'محرك توصيات التوجيه الذكي',
    quadrant: 'strategic',
    phase: 3,
    impact: 'High',
    effort: 'High',
    status: 'Proposed',
    dependencies: ['call-quality', 'attendance-tracking'],
    descEn: 'Algorithmic assessment of rep gaps (e.g. poor morning call rates) to offer action-oriented coaching advice for supervisors.',
    descAr: 'تقييم خوارزمي لنقاط ضعف المندوبين (مثل ضعف معدل زيارات الصباح) لتقديم نصائح توجيهية عملية للمشرفين.',
    businessEn: 'Systematizes training recommendations, transforming metrics into automated, personalized manager actions.',
    businessAr: 'ينظم توصيات التدريب والتوجيه، محولاً المقاييس المكتوبة إلى توجيهات آلية ومخصصة للمشرفين.',
    stepsEn: ['Create anomaly detection logic', 'Build recommendation copy generator template', 'Display coach alerts in supervisor dashboard.'],
    stepsAr: ['إنشاء منطق كشف الشذوذ والأداء الضعيف', 'بناء نموذج توليد نصوص التوصيات التلقائية', 'عرض تنبيهات التوجيه في لوحة المشرفين.']
  },
  {
    id: 'drill-down',
    titleEn: 'Drill-Down Analytics',
    titleAr: 'تحليل التفاصيل العميقة',
    quadrant: 'strategic',
    phase: 2,
    impact: 'High',
    effort: 'Medium',
    status: 'Proposed',
    dependencies: ['audit-trail'],
    descEn: 'Enable interactive drill-downs. Click on any summary aggregate to view the underlying doctor visits and survey records.',
    descAr: 'تمكين تصفح التفاصيل العميقة التفاعلي. انقر على أي رقم إجمالي لعرض تفاصيل زيارات الأطباء وسجلات الاستبيانات الأساسية.',
    businessEn: 'Enables quick verification of high-level figures without navigating back and forth through other tables.',
    businessAr: 'يمكن من التحقق السريع من الأرقام الإجمالية دون الحاجة للتنقل ذهاباً وإياباً بين الجداول الأخرى.',
    stepsEn: ['Implement state management for active node drill', 'Create pop-up detail modals', 'Perform sub-queries on specific user rows.'],
    stepsAr: ['تنفيذ إدارة الحالة للمقاطع النشطة للتعمق', 'إنشاء نوافذ تفصيلية منبثقة', 'تنفيذ استعلامات فرعية على سجلات المندوبين المحددة.']
  },
  {
    id: 'mobile-responsive',
    titleEn: 'Mobile Layout Optimization',
    titleAr: 'تحسين المظهر للهواتف الذكية',
    quadrant: 'polish',
    phase: 2,
    impact: 'Medium',
    effort: 'Medium',
    status: 'Proposed',
    dependencies: [],
    descEn: 'Enhance dashboard CSS and tables to collapse elegantly into card lists and swipeable rows on mobile screens.',
    descAr: 'تحسين ملفات التنسيق CSS والجداول لتتحول بمرونة إلى قوائم بطاقات وصفوف قابلة للسحب على شاشات الهواتف.',
    businessEn: 'Empowers reps and supervisors in the field, giving them quick, clear access to metrics on their phones.',
    businessAr: 'يمكّن المندوبين والمشرفين في الميدان من الوصول السريع والواضح لمقاييس الأداء عبر هواتفهم الشخصية.',
    stepsEn: ['Apply media query rules to table wrappers', 'Build scrollable layouts for data grids', 'Adjust header tap targets for thumbs.'],
    stepsAr: ['تطبيق قواعد الاستعلام عن الشاشة (Media Query) للجداول', 'بناء تخطيطات مرنة وقابلة للتمرير للشبكات', 'ضبط أزرار التفاعل لتناسب اللمس بالهاتف.']
  },
  {
    id: 'custom-dashboards',
    titleEn: 'Customizable Dashboards',
    titleAr: 'تخصيص لوحة التحكم والواجهات',
    quadrant: 'polish',
    phase: 3,
    impact: 'High',
    effort: 'High',
    status: 'Proposed',
    dependencies: ['quick-filters'],
    descEn: 'Allow managers to customize their layout preferences, select focus metrics, and save their configurations.',
    descAr: 'السماح للمدراء بتعديل تفضيلات تخطيط الواجهة، اختيار مؤشرات التركيز وحفظ إعداداتهم المفضلة.',
    businessEn: 'Gives individual teams the freedom to prioritize localized parameters, speeding up regular meetings.',
    businessAr: 'يمنح الفرق المختلفة حرية ترتيب الأولويات المحلية، مما يسرع الاجتماعات الدورية للمبيعات.',
    stepsEn: ['Implement drag-and-drop dashboard grid', 'Store manager UI preferences in local storage or profile DB', 'Build toggle settings panel.'],
    stepsAr: ['تنفيذ شبكة سحب وإفلات لبطاقات الواجهة', 'حفظ تفضيلات المشرفين في التخزين المحلي أو قاعدة البيانات', 'بناء لوحة إعدادات التبديل.']
  },
  {
    id: 'team-comparison',
    titleEn: 'Side-by-side Division Benchmarking',
    titleAr: 'مقارنة خطوط العمل المباشرة',
    quadrant: 'polish',
    phase: 1,
    impact: 'High',
    effort: 'Low',
    status: 'Proposed',
    dependencies: [],
    descEn: 'Compare performance side-by-side (e.g., Focus 1 vs Focus 2 vs Platinum) in unified performance panels.',
    descAr: 'مقارنة الأداء جنباً إلى جنب (مثال: Focus 1 مقابل Focus 2 مقابل Platinum) في لوحات أداء موحدة.',
    businessEn: 'Gives senior executives immediate insights into which organizational line is driving overall growth.',
    businessAr: 'يمنح القيادة التنفيذية رؤية فورية لخط العمل التنظيمي الذي يقود النمو الإجمالي.',
    stepsEn: ['Create team division grouping filters', 'Build cross-tab comparison grids', 'Incorporate split-bar comparison charts.'],
    stepsAr: ['إنشاء فلاتر تجميع أقسام الفريق', 'بناء شبكات مقارنة متعددة الأقسام', 'إدراج مخططات الأعمدة المقسمة للمقارنة.']
  },
  {
    id: 'validation-warnings',
    titleEn: 'Outlier & Verification Warnings',
    titleAr: 'تنبيهات التحقق والحالات الشاذة',
    quadrant: 'infra',
    phase: 1,
    impact: 'Medium',
    effort: 'Medium',
    status: 'Proposed',
    dependencies: [],
    descEn: 'Automatic warning badges flags for missing fields, extreme duration metrics, or overlap call locations.',
    descAr: 'شارات تحذيرية تلقائية لحقول البيانات المفقودة، أو مقاييس المدة الزمنية المبالغ فيها، أو تداخل مواقع الزيارة.',
    businessEn: 'Ensures data integrity and authenticity, preventing database errors before they impact final payroll or targets.',
    businessAr: 'يضمن سلامة البيانات وموثوقيتها، ويمنع الأخطاء التحليلية قبل تأثيرها على الرواتب أو المستهدفات.',
    stepsEn: ['Write server-side data check triggers', 'Define parameters for statistical outliers', 'Render warning alerts on summaries.'],
    stepsAr: ['كتابة شروط التحقق التلقائي للبيانات', 'تعريف معايير الحالات الإحصائية الشاذة', 'عرض شارات التحذير في جداول الملخصات.']
  },
  {
    id: 'audit-trail',
    titleEn: 'Data Update Logs & Audit Trails',
    titleAr: 'سجلات التعديل وتدقيق البيانات',
    quadrant: 'infra',
    phase: 2,
    impact: 'Medium',
    effort: 'Medium',
    status: 'Proposed',
    dependencies: [],
    descEn: 'Detailed action logger tracing who modified survey records, which administrator edited targets, and when.',
    descAr: 'سجل تفصيلي يوضح من قام بتعديل سجلات الاستبيان، وأي من المسؤولين عدل المستهدفات وتوقيت ذلك.',
    businessEn: 'Creates a transparent environment and speeds up data dispute resolutions inside the organization.',
    businessAr: 'يخلق بيئة عمل شفافة ويسرع حل النزاعات المتعلقة بالبيانات داخل المؤسسة.',
    stepsEn: ['Create database logging table', 'Bind update triggers on tables', 'Develop audit logs viewer component.'],
    stepsAr: ['إنشاء جدول سجلات التدقيق في قاعدة البيانات', 'ربط محفزات التحديث (Triggers) بالجداول', 'تطوير لوحة عرض سجلات التدقيق.']
  },
  {
    id: 'attendance-tracking',
    titleEn: 'Attendance & Sick Leave Tracker',
    titleAr: 'متابعة الحضور والإجازات المرضية',
    quadrant: 'infra',
    phase: 2,
    impact: 'Medium',
    effort: 'Medium',
    status: 'Proposed',
    dependencies: [],
    descEn: 'Log days out of field (e.g. sick leaves, training workshops) to explain performance gap occurrences in summaries.',
    descAr: 'تسجيل أيام الغياب عن الميدان (مثل الإجازات المرضية، ورش العمل) لتوضيح ثغرات الأداء الظاهرة في الملخصات.',
    businessEn: 'Avoids penalizing reps for scheduled training or approved leaves, enabling fair and accurate evaluations.',
    businessAr: 'يتجنب معاقبة المندوبين بسبب التدريبات المجدولة أو الإجازات المعتمدة، مما يضمن تقييمات عادلة ودقيقة.',
    stepsEn: ['Create out-of-office submission forms', 'Factor leave days into call rate averages', 'Overlay leave labels on calendars.'],
    stepsAr: ['إنشاء نماذج تقديم طلبات الغياب', 'إدخال أيام الإجازة في حساب معدل الزيارات', 'تراكب ملصقات الإجازة في التقويم السنوي.']
  }
];

export default function EnhancementRoadmap({ lang, rtl }) {
  const [subTab, setSubTab] = useState('matrix'); // 'matrix' | 'timeline' | 'dependencies'
  const [selectedItem, setSelectedItem] = useState(null);
  const [dependencyHover, setDependencyHover] = useState(null);

  // Level Slicers State
  const [filterImpact, setFilterImpact] = useState('all'); // 'all' | 'High' | 'Medium'
  const [filterEffort, setFilterEffort] = useState('all'); // 'all' | 'High' | 'Medium' | 'Low'
  const [filterPhase, setFilterPhase] = useState('all');   // 'all' | '1' | '2' | '3'

  // Match check helper
  const isMatch = (item) => {
    if (filterImpact !== 'all' && item.impact !== filterImpact) return false;
    if (filterEffort !== 'all' && item.effort !== filterEffort) return false;
    if (filterPhase !== 'all' && String(item.phase) !== String(filterPhase)) return false;
    return true;
  };

  const hasActiveFilters = filterImpact !== 'all' || filterEffort !== 'all' || filterPhase !== 'all';

  const clearFilters = () => {
    setFilterImpact('all');
    setFilterEffort('all');
    setFilterPhase('all');
  };

  // Localization labels
  const UI = {
    en: {
      title: 'Enhancement Roadmap',
      subtitle: 'Explore the technical execution strategy and dependency mappings for upcoming CRM upgrades.',
      matrix: 'Impact vs Effort Matrix',
      timeline: 'Implementation Timeline',
      dependencies: 'Dependencies Map',
      back: 'Back to Roadmap',
      quadrants: {
        'quick-wins': '⚡ Quick Wins (High Impact, Low Effort)',
        'strategic': '🎯 Strategic Investments (High Impact, High Effort)',
        'polish': '✨ Polish & Refinement (Low Impact, Low Effort)',
        'infra': '🔧 Infrastructure (Low Impact, High Effort)'
      },
      phase: 'Phase',
      effort: 'Effort',
      impact: 'Impact',
      status: 'Status',
      businessImpact: 'Business Impact & ROI',
      steps: 'Key Implementation Steps',
      preReq: 'Prerequisites',
      downstream: 'Enables Downstream',
      none: 'None',
      clickPrompt: 'Click on any enhancement to view execution details',
      timelinePhase1: 'Phase 1: Foundation & Quick Wins (Months 1-2)',
      timelinePhase2: 'Phase 2: Insights & Mobility (Months 3-4)',
      timelinePhase3: 'Phase 3: AI & Customization (Months 5-6+)',
      close: 'Close Panel',
      slicersTitle: 'Roadmap Slicers',
      clearAll: 'Clear Filters',
      levels: {
        all: 'All Levels',
        High: 'High',
        Medium: 'Medium',
        Low: 'Low',
        1: 'Phase 1',
        2: 'Phase 2',
        3: 'Phase 3'
      }
    },
    ar: {
      title: 'خارطة طريق التحسينات والتطوير',
      subtitle: 'استكشف استراتيجية التنفيذ الفني وارتباطات الاعتمادية للتحديثات القادمة لمنصة CRM.',
      matrix: 'مصفوفة الأثر مقابل الجهد',
      timeline: 'الجدول الزمني للتنفيذ',
      dependencies: 'خريطة الاعتماديات',
      back: 'العودة للخارطة',
      quadrants: {
        'quick-wins': '⚡ مكاسب سريعة (أثر عالٍ، جهد منخفض)',
        'strategic': '🎯 استثمارات استراتيجية (أثر عالٍ، جهد مرتفع)',
        'polish': '✨ صقل وتحسين (أثر منخفض، جهد منخفض)',
        'infra': '🔧 البنية التحتية (أثر منخفض، جهد مرتفع)'
      },
      phase: 'الفترة',
      effort: 'الجهد المطلوب',
      impact: 'الأثر المتوقع',
      status: 'الحالة',
      businessImpact: 'العائد والأثر على الأعمال',
      steps: 'خطوات التنفيذ الأساسية',
      preReq: 'المتطلبات السابقة',
      downstream: 'يمكّن الميزات التالية',
      none: 'لا يوجد',
      clickPrompt: 'انقر على أي بطاقة تحسين لعرض تفاصيل التنفيذ والفائدة',
      timelinePhase1: 'المرحلة الأولى: التأسيس والمكاسب السريعة (الشهر 1-2)',
      timelinePhase2: 'المرحلة الثانية: التحليلات والهواتف (الشهر 3-4)',
      timelinePhase3: 'المرحلة الثالثة: الذكاء الاصطناعي والتخصيص (الشهر 5-6+)',
      close: 'إغلاق اللوحة',
      slicersTitle: 'فلاتر خارطة الطريق',
      clearAll: 'إعادة تعيين الفلاتر',
      levels: {
        all: 'الكل',
        High: 'مرتفع',
        Medium: 'متوسط',
        Low: 'منخفض',
        1: 'المرحلة الأولى',
        2: 'المرحلة الثانية',
        3: 'المرحلة الثالثة'
      }
    }
  }[lang] || { en: {} };

  // Calculate dependencies for displaying in details panel
  const getDependencyNames = (ids) => {
    if (!ids || ids.length === 0) return UI.none;
    return ids.map(id => {
      const match = ROADMAP_ITEMS.find(item => item.id === id);
      return match ? (lang === 'ar' ? match.titleAr : match.titleEn) : id;
    }).join(', ');
  };

  const getDownstreamFeatures = (id) => {
    const list = ROADMAP_ITEMS.filter(item => item.dependencies.includes(id));
    if (list.length === 0) return UI.none;
    return list.map(item => (lang === 'ar' ? item.titleAr : item.titleEn)).join(', ');
  };

  // Node connection coordinate points for dependencies drawing (based on visual indices)
  const nodes = [
    { id: 'progress-cards', x: 80, y: 70 },
    { id: 'benchmarking', x: 270, y: 70 },
    { id: 'quick-filters', x: 80, y: 160 },
    { id: 'custom-dashboards', x: 440, y: 160 },
    { id: 'validation-warnings', x: 80, y: 250 },
    { id: 'territory-map', x: 440, y: 250 },
    { id: 'audit-trail', x: 80, y: 340 },
    { id: 'drill-down', x: 270, y: 340 },
    { id: 'call-quality', x: 270, y: 430 },
    { id: 'attendance-tracking', x: 270, y: 520 },
    { id: 'coaching-engine', x: 600, y: 475 },
    // Standalone items positioned on side lanes
    { id: 'trend-lines', x: 740, y: 70 },
    { id: 'export-templates', x: 740, y: 160 },
    { id: 'mobile-responsive', x: 740, y: 250 },
    { id: 'team-comparison', x: 740, y: 340 },
  ];

  return (
    <div className={`roadmap-container ${rtl ? 'rtl' : 'ltr'}`}>
      <div className="roadmap-header-section">
        <h2>{UI.title}</h2>
        <p className="roadmap-subtitle">{UI.subtitle}</p>
      </div>

      {/* ROADMAP VIEWS NAVIGATION */}
      <div className="roadmap-navigation">
        <button 
          className={`roadmap-nav-btn ${subTab === 'matrix' ? 'active' : ''}`}
          onClick={() => setSubTab('matrix')}
        >
          {UI.matrix}
        </button>
        <button 
          className={`roadmap-nav-btn ${subTab === 'timeline' ? 'active' : ''}`}
          onClick={() => setSubTab('timeline')}
        >
          {UI.timeline}
        </button>
        <button 
          className={`roadmap-nav-btn ${subTab === 'dependencies' ? 'active' : ''}`}
          onClick={() => setSubTab('dependencies')}
        >
          {UI.dependencies}
        </button>
      </div>

      {/* LEVEL SLICERS BAR */}
      <div className="roadmap-slicers-bar">
        <div className="slicers-bar-hdr">
          <span className="slicers-bar-title">🔍 {UI.slicersTitle}</span>
          {hasActiveFilters && (
            <button className="slicers-clear-btn" onClick={clearFilters}>
              {UI.clearAll}
            </button>
          )}
        </div>

        <div className="slicers-grid">
          {/* Impact Level Slicer */}
          <div className="slicer-group">
            <span className="slicer-lbl">{UI.impact}</span>
            <div className="slicer-options">
              {['all', 'High', 'Medium'].map(lvl => (
                <button
                  key={lvl}
                  className={`slicer-btn ${filterImpact === lvl ? 'active' : ''}`}
                  onClick={() => setFilterImpact(lvl)}
                >
                  {UI.levels[lvl]}
                </button>
              ))}
            </div>
          </div>

          {/* Effort Level Slicer */}
          <div className="slicer-group">
            <span className="slicer-lbl">{UI.effort}</span>
            <div className="slicer-options">
              {['all', 'High', 'Medium', 'Low'].map(lvl => (
                <button
                  key={lvl}
                  className={`slicer-btn ${filterEffort === lvl ? 'active' : ''}`}
                  onClick={() => setFilterEffort(lvl)}
                >
                  {UI.levels[lvl]}
                </button>
              ))}
            </div>
          </div>

          {/* Phase Slicer */}
          <div className="slicer-group">
            <span className="slicer-lbl">{UI.phase}</span>
            <div className="slicer-options">
              {['all', '1', '2', '3'].map(lvl => (
                <button
                  key={lvl}
                  className={`slicer-btn ${filterPhase === lvl ? 'active' : ''}`}
                  onClick={() => setFilterPhase(lvl)}
                >
                  {UI.levels[lvl]}
                </button>
              ))}
            </div>
          </div>
        </div>
      </div>

      <div className="roadmap-content-area">
        {/* View 1: Impact-Effort 2x2 Matrix */}
        {subTab === 'matrix' && (
          <div className="matrix-grid-view">
            {['quick-wins', 'strategic', 'polish', 'infra'].map(q => (
              <div key={q} className={`matrix-quadrant ${q}`}>
                <div className="quadrant-header">{UI.quadrants[q]}</div>
                <div className="quadrant-body">
                  {ROADMAP_ITEMS.filter(item => item.quadrant === q).map(item => {
                    const matchState = isMatch(item);
                    return (
                      <div 
                        key={item.id} 
                        className={`roadmap-card ${selectedItem?.id === item.id ? 'selected' : ''} ${!matchState ? 'unmatched' : ''}`}
                        onClick={() => setSelectedItem(item)}
                      >
                        <div className="card-top">
                          <span className={`phase-badge phase-${item.phase}`}>
                            {UI.phase} {item.phase}
                          </span>
                          <span className={`status-badge status-${item.status.toLowerCase()}`}>
                            {item.status}
                          </span>
                        </div>
                        <h4 className="card-title">{lang === 'ar' ? item.titleAr : item.titleEn}</h4>
                        <p className="card-summary">
                          {lang === 'ar' ? item.descAr.slice(0, 75) : item.descEn.slice(0, 75)}...
                        </p>
                        <div className="card-metrics-footer">
                          <span>⚡ {UI.impact}: <strong>{item.impact}</strong></span>
                          <span>🔧 {UI.effort}: <strong>{item.effort}</strong></span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        )}

        {/* View 2: Implementation Timeline */}
        {subTab === 'timeline' && (
          <div className="timeline-gantt-view">
            {[1, 2, 3].map(phaseNum => (
              <div key={phaseNum} className={`timeline-phase-lane phase-lane-${phaseNum}`}>
                <div className="timeline-phase-title">
                  {phaseNum === 1 ? UI.timelinePhase1 : phaseNum === 2 ? UI.timelinePhase2 : UI.timelinePhase3}
                </div>
                <div className="timeline-phase-cards">
                  {ROADMAP_ITEMS.filter(item => item.phase === phaseNum).map(item => {
                    const matchState = isMatch(item);
                    return (
                      <div 
                        key={item.id}
                        className={`timeline-item-row ${selectedItem?.id === item.id ? 'selected' : ''} ${!matchState ? 'unmatched' : ''}`}
                        onClick={() => setSelectedItem(item)}
                      >
                        <div className="timeline-item-title-col">
                          <strong>{lang === 'ar' ? item.titleAr : item.titleEn}</strong>
                          <span className={`quadrant-tag tag-${item.quadrant}`}>
                            {item.quadrant.toUpperCase().replace('-', ' ')}
                          </span>
                        </div>
                        <div className="timeline-progress-bar-container">
                          <div className={`timeline-progress-fill phase-fill-${item.phase}`}>
                            {item.effort} Effort
                          </div>
                        </div>
                        <div className="timeline-item-status-col">
                          <span className={`status-badge status-${item.status.toLowerCase()}`}>
                            {item.status}
                          </span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        )}

        {/* View 3: SVG Dependencies Flowchart */}
        {subTab === 'dependencies' && (
          <div className="dependencies-svg-view">
            <div className="dependencies-legend">
              <span className="legend-item"><span className="legend-circle met"/> Foundation / Complete</span>
              <span className="legend-item"><span className="legend-circle dependency"/> High Dependency Links</span>
              <span className="legend-item"><span className="legend-circle independent"/> Standalone / Auxiliary Modules</span>
            </div>
            
            <div className="svg-canvas-container">
              <svg className="dependencies-svg" width="920" height="600">
                <defs>
                  <marker id="arrow" viewBox="0 0 10 10" refX="22" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse">
                    <path d="M 0 0 L 10 5 L 0 10 z" fill="#888" />
                  </marker>
                  <marker id="arrow-active" viewBox="0 0 10 10" refX="22" refY="5" markerWidth="8" markerHeight="8" orient="auto-start-reverse">
                    <path d="M 0 0 L 10 5 L 0 10 z" fill="#3b82f6" />
                  </marker>
                </defs>

                {/* Draw connection paths */}
                {ROADMAP_ITEMS.map(item => {
                  const nodeObj = nodes.find(n => n.id === item.id);
                  if (!nodeObj) return null;

                  return item.dependencies.map(depId => {
                    const depNode = nodes.find(n => n.id === depId);
                    if (!depNode) return null;

                    const isHighlighted = dependencyHover === item.id || dependencyHover === depId || selectedItem?.id === item.id || selectedItem?.id === depId;
                    const pathMatching = isMatch(item) && isMatch(ROADMAP_ITEMS.find(r => r.id === depId));

                    return (
                      <path
                        key={`${depId}-${item.id}`}
                        d={`M ${depNode.x + 85} ${depNode.y + 25} C ${(depNode.x + nodeObj.x) / 2 + 85} ${depNode.y + 25}, ${(depNode.x + nodeObj.x) / 2 + 85} ${nodeObj.y + 25}, ${nodeObj.x + 85} ${nodeObj.y + 25}`}
                        fill="none"
                        stroke={isHighlighted ? '#3b82f6' : 'rgba(255, 255, 255, 0.15)'}
                        strokeWidth={isHighlighted ? 3 : 1.5}
                        markerEnd={isHighlighted ? "url(#arrow-active)" : "url(#arrow)"}
                        className={`connection-path ${!pathMatching ? 'unmatched-path' : ''}`}
                      />
                    );
                  });
                })}

                {/* Render nodes inside foreignObject to allow rich CSS cards */}
                {nodes.map(node => {
                  const item = ROADMAP_ITEMS.find(i => i.id === node.id);
                  if (!item) return null;

                  const isHoveredOrSelected = selectedItem?.id === item.id || dependencyHover === item.id;
                  const hasDeps = item.dependencies.length > 0;
                  const enablesOthers = ROADMAP_ITEMS.some(i => i.dependencies.includes(item.id));
                  const matchState = isMatch(item);

                  return (
                    <foreignObject
                      key={node.id}
                      x={node.x}
                      y={node.y}
                      width="170"
                      height="54"
                      onMouseEnter={() => setDependencyHover(item.id)}
                      onMouseLeave={() => setDependencyHover(null)}
                      onClick={() => setSelectedItem(item)}
                    >
                      <div className={`dep-node-card ${isHoveredOrSelected ? 'active' : ''} ${hasDeps ? 'has-deps' : ''} ${enablesOthers ? 'enabler' : ''} ${!matchState ? 'unmatched' : ''}`}>
                        <div className="dep-node-header">
                          <span className={`quadrant-dot dot-${item.quadrant}`} />
                          <span className="dep-node-phase">P{item.phase}</span>
                        </div>
                        <div className="dep-node-title" title={lang === 'ar' ? item.titleAr : item.titleEn}>
                          {lang === 'ar' ? item.titleAr : item.titleEn}
                        </div>
                      </div>
                    </foreignObject>
                  );
                })}
              </svg>
            </div>
          </div>
        )}
      </div>

      {/* DETAIL SIDE PANEL (DRAWER) */}
      {selectedItem && (
        <>
          <div className="roadmap-drawer-backdrop" onClick={() => setSelectedItem(null)} />
          <div className={`roadmap-drawer ${rtl ? 'rtl' : 'ltr'}`}>
            <div className="drawer-header">
              <h3>{lang === 'ar' ? selectedItem.titleAr : selectedItem.titleEn}</h3>
              <button className="drawer-close-btn" onClick={() => setSelectedItem(null)}>✕</button>
            </div>

            <div className="drawer-body">
              <div className="drawer-badges">
                <span className={`quadrant-tag tag-${selectedItem.quadrant}`}>
                  {UI.quadrants[selectedItem.quadrant] || selectedItem.quadrant}
                </span>
                <span className={`phase-badge phase-${selectedItem.phase}`}>
                  {UI.phase} {selectedItem.phase}
                </span>
                <span className={`status-badge status-${selectedItem.status.toLowerCase()}`}>
                  {selectedItem.status}
                </span>
              </div>

              <div className="drawer-section">
                <h4>📋 {lang === 'ar' ? 'الوصف' : 'Description'}</h4>
                <p className="drawer-text-highlight">
                  {lang === 'ar' ? selectedItem.descAr : selectedItem.descEn}
                </p>
              </div>

              <div className="drawer-section">
                <h4>🎯 {UI.businessImpact}</h4>
                <p className="drawer-text">
                  {lang === 'ar' ? selectedItem.businessAr : selectedItem.businessEn}
                </p>
              </div>

              <div className="drawer-section">
                <h4>🛠️ {UI.steps}</h4>
                <ul className="drawer-steps-list">
                  {(lang === 'ar' ? selectedItem.stepsAr : selectedItem.stepsEn).map((step, idx) => (
                    <li key={idx}>
                      <span className="step-number">{idx + 1}</span>
                      <span className="step-label">{step}</span>
                    </li>
                  ))}
                </ul>
              </div>

              <div className="drawer-divider" />

              <div className="drawer-meta-grid">
                <div className="meta-box">
                  <span className="meta-lbl">{UI.effort}</span>
                  <span className="meta-val val-effort">{selectedItem.effort}</span>
                </div>
                <div className="meta-box">
                  <span className="meta-lbl">{UI.impact}</span>
                  <span className="meta-val val-impact">{selectedItem.impact}</span>
                </div>
              </div>

              <div className="drawer-section">
                <div className="dep-flow-info">
                  <strong>⛓️ {UI.preReq}:</strong>
                  <div className="dep-flow-names">{getDependencyNames(selectedItem.dependencies)}</div>
                </div>
                <div className="dep-flow-info" style={{ marginTop: '12px' }}>
                  <strong>🚀 {UI.downstream}:</strong>
                  <div className="dep-flow-names">{getDownstreamFeatures(selectedItem.id)}</div>
                </div>
              </div>
            </div>

            <div className="drawer-footer">
              <button className="hbtn hbtn-primary drawer-close-action" onClick={() => setSelectedItem(null)}>
                {UI.close}
              </button>
            </div>
          </div>
        </>
      )}
    </div>
  );
}
