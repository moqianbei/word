//生成目录索引列表
function GenerateContentList() {
    var jquery_h1_list = $('#mainContent h2');
    if (jquery_h1_list.length == 0) { return; }
    if ($('#allNav').length == 0) { return; }

    var content = '<a name="_labelTop"></a>';
    content += '<div id="navCategory">';
    content += '<p style="font-size:1.5rem;font-weight:bold;">目录</p>';
    // 一级目录 start
    content += '<ul class="first_class_ul">';

    for (var i = 0; i < jquery_h1_list.length; i++) {
        var go_to_top = '<div><a name="_label' + i + '"></a></div>';
        $(jquery_h1_list[i]).before(go_to_top);

        // 一级目录的一条
        var li_content = '<li><a href="#_label' + i + '" rel="external nofollow" rel="external nofollow" rel="external nofollow" rel="external nofollow" >' + $(jquery_h1_list[i]).text() + '</a></li>';

        var nextH1Index = i + 1;
        if (nextH1Index == jquery_h1_list.length) { nextH1Index = 0; }
        var jquery_h2_list = $(jquery_h1_list[i]).nextUntil(jquery_h1_list[nextH1Index], "h3");
        // 二级目录 start
        if (jquery_h2_list.length > 0) {
            //li_content +='<ul style="list-style-type:none; text-align: left; margin:2px 2px;">';
            li_content += '<ul class="second_class_ul">';
        }
        for (var j = 0; j < jquery_h2_list.length; j++) {
            var go_to_top2 = '<div><a name="_lab2_' + i + '_' + j + '"></a></div>';
            $(jquery_h2_list[j]).before(go_to_top2);
            // 二级目录的一条
            li_content += '<li><a href="#_lab2_' + i + '_' + j + '" rel="external nofollow" >' + $(jquery_h2_list[j]).text() + '</a></li>';

            var nextH2Index = j + 1;
            var next;
            if (nextH2Index == jquery_h2_list.length) {
                if (i + 1 == jquery_h1_list.length) {
                    next = jquery_h1_list[0];
                } else {
                    next = jquery_h1_list[i + 1];
                }
            } else {
                next = jquery_h2_list[nextH2Index];
            }
            var jquery_h3_list = $(jquery_h2_list[j]).nextUntil(next, "h4");

            // 三级目录 start
            if (jquery_h3_list.length > 0) {
                li_content += '<ul class="third_class_ul">';
            }

            for (var k = 0; k < jquery_h3_list.length; k++) {
                var go_to_third_Content = '<div style="text-align: right"><a name="_label3_' + i + '_' + j + '_' + k + '"></a></div>';
                $(jquery_h3_list[k]).before(go_to_third_Content);
                // 三级目录的一条
                li_content += '<li><a href="#_label3_' + i + '_' + j + '_' + k + '" rel="external nofollow" >' + $(jquery_h3_list[k]).text() + '</a></li>';
            }

            if (jquery_h3_list.length > 0) {
                li_content += '</ul>';
            }
            li_content += '</li>';
            // 三级目录 end
        }
        if (jquery_h2_list.length > 0) {
            li_content += '</ul>';
        }
        li_content += '</li>';
        // 二级目录 end

        content += li_content;
    }
    // 一级目录 end
    content += '</ul>';
    content += '</div>';

    $($('#allNav')[0]).prepend(content);
}

GenerateContentList();

// 自动更新进度条

var prg = $("progress");

for (var p = 0; p < prg.length; p++) {

    prg[p].value = prg[p].offsetTop;
    prg[p].max = document.body.clientHeight;
}
