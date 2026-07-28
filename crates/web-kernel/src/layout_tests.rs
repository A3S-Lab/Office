use super::*;

fn request(blocks: Vec<LayoutBlock>) -> LayoutRequest {
    LayoutRequest {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: "layout".into(),
        request_id: 7,
        revision: 12,
        document_revision: 4,
        start_page_index: 0,
        page: LayoutPageMetrics {
            width: 794.0,
            height: 1123.0,
            margin_top: 80.0,
            margin_right: 80.0,
            margin_bottom: 80.0,
            margin_left: 80.0,
            header_height: 0.0,
            footer_height: 0.0,
            page_gap: 24.0,
        },
        blocks,
    }
}

fn block(id: &str, height: f64) -> LayoutBlock {
    LayoutBlock {
        id: id.into(),
        height,
        break_before: false,
        break_after: false,
        keep_together: false,
        keep_with_next: false,
        flow_id: None,
        flow_index: None,
        flow_count: None,
        minimum_fragments_per_page: None,
        repeat_header_count: None,
        repeat_header_height: None,
    }
}

fn presentation_request(alignment: PresentationAlignment) -> PresentationGeometryRequest {
    PresentationGeometryRequest {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: "presentationGeometry".into(),
        request_id: 9,
        revision: 3,
        document_revision: 2,
        operation: PresentationGeometryOperation::AlignToSlide { alignment },
        elements: vec![PresentationGeometryElement {
            id: "title".into(),
            x: 17.0,
            y: 23.0,
            width: 40.0,
            height: 20.0,
        }],
    }
}

#[test]
fn aligns_presentation_elements_to_the_slide() {
    let center =
        resolve_presentation_geometry(&presentation_request(PresentationAlignment::Center))
            .expect("center alignment");
    let bottom =
        resolve_presentation_geometry(&presentation_request(PresentationAlignment::Bottom))
            .expect("bottom alignment");

    assert_eq!(center.elements[0].x, 30.0);
    assert_eq!(center.elements[0].y, 23.0);
    assert!(center.guides.is_empty());
    assert_eq!(bottom.elements[0].x, 17.0);
    assert_eq!(bottom.elements[0].y, 80.0);
    assert_eq!(bottom.document_revision, 2);
}

#[test]
fn snaps_presentation_moves_and_resizes_with_stable_guides() {
    let mut request = presentation_request(PresentationAlignment::Center);
    request.elements = vec![
        PresentationGeometryElement {
            id: "moving".into(),
            x: 29.4,
            y: 10.0,
            width: 40.0,
            height: 20.0,
        },
        PresentationGeometryElement {
            id: "target".into(),
            x: 72.0,
            y: 61.0,
            width: 18.0,
            height: 12.0,
        },
    ];
    request.operation = PresentationGeometryOperation::SnapElement {
        moving_element_id: "moving".into(),
        mode: PresentationTransformMode::Move,
        threshold_x: 1.0,
        threshold_y: 1.0,
    };

    let moved = resolve_presentation_geometry(&request).expect("move snap");
    assert_eq!(moved.elements[0].x, 30.0);
    assert_eq!(
        moved.guides,
        vec![PresentationSnapGuide {
            axis: PresentationSnapGuideAxis::X,
            position: 50.0,
            source: PresentationSnapGuideSource::Slide,
            target_id: None,
        }]
    );

    request.elements[0].x = 10.0;
    request.elements[0].width = 39.4;
    request.operation = PresentationGeometryOperation::SnapElement {
        moving_element_id: "moving".into(),
        mode: PresentationTransformMode::Resize,
        threshold_x: 1.0,
        threshold_y: 0.0,
    };
    let resized = resolve_presentation_geometry(&request).expect("resize snap");
    assert_eq!(resized.elements[0].width, 40.0);
    assert_eq!(resized.guides[0].position, 50.0);

    request.elements[0].x = 69.4;
    request.elements[0].width = 2.0;
    request.operation = PresentationGeometryOperation::SnapElement {
        moving_element_id: "moving".into(),
        mode: PresentationTransformMode::Move,
        threshold_x: 1.0,
        threshold_y: 0.0,
    };
    let sibling = resolve_presentation_geometry(&request).expect("sibling snap");
    assert_eq!(sibling.elements[0].x, 70.0);
    assert_eq!(
        sibling.guides,
        vec![PresentationSnapGuide {
            axis: PresentationSnapGuideAxis::X,
            position: 72.0,
            source: PresentationSnapGuideSource::Element,
            target_id: Some("target".into()),
        }]
    );

    request.elements[0].x = 0.0;
    request.elements[0].width = 1.0;
    request.operation = PresentationGeometryOperation::SnapElement {
        moving_element_id: "moving".into(),
        mode: PresentationTransformMode::Resize,
        threshold_x: 1.0,
        threshold_y: 0.0,
    };
    let preserved = resolve_presentation_geometry(&request).expect("non-collapsing resize");
    assert_eq!(preserved.elements[0].width, 1.0);
    assert!(preserved.guides.is_empty());
}

#[test]
fn paginates_blocks_and_reports_visual_spacers() {
    let result = layout_document(&request(vec![
        block("one", 500.0),
        block("two", 500.0),
        block("three", 120.0),
    ]))
    .expect("layout");

    assert_eq!(result.pages.len(), 2);
    assert_eq!(result.document_revision, 4);
    assert_eq!(result.pages[0].placements.len(), 1);
    assert_eq!(result.pages[1].placements.len(), 2);
    assert_eq!(result.breaks[0].before_block_id, "two");
    assert_eq!(result.breaks[0].spacer_height, 647.0);
}

#[test]
fn keeps_page_chrome_inside_the_physical_margins() {
    let mut input = request(vec![block("body", 150.0), block("tail", 20.0)]);
    input.page = LayoutPageMetrics {
        width: 300.0,
        height: 200.0,
        margin_top: 20.0,
        margin_right: 20.0,
        margin_bottom: 20.0,
        margin_left: 20.0,
        header_height: 20.0,
        footer_height: 20.0,
        page_gap: 30.0,
    };

    let result = layout_document(&input).expect("layout");

    assert_eq!(result.pages.len(), 2);
    assert_eq!(result.pages[0].available_height, 160.0);
    assert!(!result.pages[0].placements[0].overflow);
    assert_eq!(result.breaks[0].before_block_id, "tail");
    assert_eq!(result.breaks[0].remaining_body_height, 10.0);
    assert_eq!(result.breaks[0].spacer_height, 80.0);
}

#[test]
fn keeps_a_heading_with_the_following_block() {
    let mut heading = block("heading", 40.0);
    heading.keep_with_next = true;
    let result = layout_document(&request(vec![
        block("intro", 850.0),
        heading,
        block("paragraph", 100.0),
    ]))
    .expect("layout");

    assert_eq!(result.pages.len(), 2);
    assert_eq!(result.pages[1].placements[0].block_id, "heading");
    assert_eq!(result.pages[1].placements[1].block_id, "paragraph");
}

#[test]
fn lays_out_a_suffix_with_absolute_page_indices() {
    let mut input = request(vec![
        block("three", 500.0),
        block("four", 500.0),
        block("five", 120.0),
    ]);
    input.start_page_index = 2;

    let result = layout_document(&input).expect("layout");

    assert_eq!(result.start_page_index, 2);
    assert_eq!(
        result
            .pages
            .iter()
            .map(|page| page.index)
            .collect::<Vec<_>>(),
        vec![2, 3]
    );
    assert_eq!(result.breaks[0].page_index, 3);
    assert_eq!(result.breaks[0].before_block_id, "four");
}

#[test]
fn keeps_minimum_fragments_on_both_sides_of_a_page_break() {
    let mut blocks = vec![block("intro", 843.0)];
    blocks.extend((0..3).map(|index| LayoutBlock {
        id: format!("paragraph-line-{index}"),
        height: 60.0,
        flow_id: Some("paragraph".into()),
        flow_index: Some(index),
        flow_count: Some(3),
        minimum_fragments_per_page: Some(2),
        ..block("line", 60.0)
    }));

    let result = layout_document(&request(blocks)).expect("layout");

    assert_eq!(result.pages.len(), 2);
    assert_eq!(
        result.pages[0]
            .placements
            .iter()
            .map(|placement| placement.block_id.as_str())
            .collect::<Vec<_>>(),
        vec!["intro"]
    );
    assert_eq!(
        result.pages[1]
            .placements
            .iter()
            .map(|placement| placement.block_id.as_str())
            .collect::<Vec<_>>(),
        vec!["paragraph-line-0", "paragraph-line-1", "paragraph-line-2"]
    );
}

#[test]
fn keeps_the_end_of_a_flow_with_the_next_flow() {
    let mut blocks = vec![block("intro", 800.0)];
    blocks.extend((0..2).map(|index| LayoutBlock {
        id: format!("current-line-{index}"),
        height: 40.0,
        keep_with_next: index == 1,
        flow_id: Some("current".into()),
        flow_index: Some(index),
        flow_count: Some(2),
        minimum_fragments_per_page: Some(1),
        ..block("current-line", 40.0)
    }));
    blocks.extend((0..2).map(|index| LayoutBlock {
        id: format!("next-line-{index}"),
        height: 50.0,
        flow_id: Some("next".into()),
        flow_index: Some(index),
        flow_count: Some(2),
        minimum_fragments_per_page: Some(2),
        ..block("next-line", 50.0)
    }));

    let result = layout_document(&request(blocks)).expect("layout");

    assert_eq!(
        result.pages[0]
            .placements
            .iter()
            .map(|placement| placement.block_id.as_str())
            .collect::<Vec<_>>(),
        vec!["intro"]
    );
    assert_eq!(
        result.pages[1]
            .placements
            .iter()
            .map(|placement| placement.block_id.as_str())
            .collect::<Vec<_>>(),
        vec![
            "current-line-0",
            "current-line-1",
            "next-line-0",
            "next-line-1"
        ]
    );
}

#[test]
fn reserves_repeated_table_headers_on_continuation_pages() {
    let mut input = request(vec![block("intro", 20.0)]);
    input.page.height = 100.0;
    input.page.margin_top = 0.0;
    input.page.margin_bottom = 0.0;
    let rows = [
        ("table-header", 20.0),
        ("table-row-1", 40.0),
        ("table-row-2", 40.0),
        ("table-row-3", 40.0),
    ];
    input.blocks.extend(
        rows.into_iter()
            .enumerate()
            .map(|(flow_index, (id, height))| LayoutBlock {
                id: id.into(),
                height,
                flow_id: Some("table".into()),
                flow_index: Some(flow_index as u32),
                flow_count: Some(4),
                minimum_fragments_per_page: Some(1),
                repeat_header_count: Some(1),
                repeat_header_height: Some(20.0),
                ..block(id, height)
            }),
    );

    let result = layout_document(&input).expect("table layout");

    assert_eq!(result.pages.len(), 2);
    assert_eq!(
        result.pages[0]
            .placements
            .iter()
            .map(|placement| placement.block_id.as_str())
            .collect::<Vec<_>>(),
        vec!["intro", "table-header", "table-row-1"]
    );
    assert_eq!(result.pages[1].used_height, 100.0);
    assert_eq!(result.pages[1].placements[0].block_id, "table-row-2");
    assert_eq!(result.pages[1].placements[0].y, 20.0);
    assert_eq!(result.pages[1].placements[1].y, 60.0);
    assert_eq!(result.breaks[0].before_block_id, "table-row-2");
}

#[test]
fn keeps_repeated_table_headers_with_the_first_body_row() {
    let mut input = request(vec![block("intro", 60.0)]);
    input.page.height = 100.0;
    input.page.margin_top = 0.0;
    input.page.margin_bottom = 0.0;
    let rows = [
        ("table-header", 20.0),
        ("table-row-1", 30.0),
        ("table-row-2", 30.0),
    ];
    input.blocks.extend(
        rows.into_iter()
            .enumerate()
            .map(|(flow_index, (id, height))| LayoutBlock {
                id: id.into(),
                height,
                flow_id: Some("table".into()),
                flow_index: Some(flow_index as u32),
                flow_count: Some(3),
                minimum_fragments_per_page: Some(1),
                repeat_header_count: Some(1),
                repeat_header_height: Some(20.0),
                ..block(id, height)
            }),
    );

    let result = layout_document(&input).expect("table layout");

    assert_eq!(
        result.pages[0]
            .placements
            .iter()
            .map(|placement| placement.block_id.as_str())
            .collect::<Vec<_>>(),
        vec!["intro"]
    );
    assert_eq!(
        result.pages[1]
            .placements
            .iter()
            .map(|placement| placement.block_id.as_str())
            .collect::<Vec<_>>(),
        vec!["table-header", "table-row-1", "table-row-2"]
    );
    assert_eq!(result.pages[1].placements[0].y, 0.0);
}

#[test]
fn rejects_incomplete_repeated_table_header_metadata() {
    let fragments = vec![
        LayoutBlock {
            id: "table-header".into(),
            flow_id: Some("table".into()),
            flow_index: Some(0),
            flow_count: Some(2),
            minimum_fragments_per_page: Some(1),
            repeat_header_count: Some(1),
            ..block("table-header", 20.0)
        },
        LayoutBlock {
            id: "table-row".into(),
            flow_id: Some("table".into()),
            flow_index: Some(1),
            flow_count: Some(2),
            minimum_fragments_per_page: Some(1),
            repeat_header_count: Some(1),
            ..block("table-row", 30.0)
        },
    ];

    let error = layout_document(&request(fragments)).expect_err("invalid repeated table header");
    assert_eq!(
        error.code,
        "office.kernel.repeat_header_metadata_incomplete"
    );
}

#[test]
fn rejects_non_consecutive_flow_fragments() {
    let fragments = vec![
        LayoutBlock {
            id: "paragraph-line-0".into(),
            flow_id: Some("paragraph".into()),
            flow_index: Some(0),
            flow_count: Some(2),
            minimum_fragments_per_page: Some(1),
            ..block("line-0", 20.0)
        },
        LayoutBlock {
            id: "paragraph-line-1".into(),
            flow_id: Some("paragraph".into()),
            flow_index: Some(0),
            flow_count: Some(2),
            minimum_fragments_per_page: Some(1),
            ..block("line-1", 20.0)
        },
    ];

    let error = layout_document(&request(fragments)).expect_err("invalid flow");
    assert_eq!(error.code, "office.kernel.flow_sequence_invalid");
}

#[test]
fn rejects_page_breaks_inside_a_flow() {
    let fragments = vec![
        LayoutBlock {
            id: "paragraph-line-0".into(),
            break_after: true,
            flow_id: Some("paragraph".into()),
            flow_index: Some(0),
            flow_count: Some(2),
            minimum_fragments_per_page: Some(1),
            ..block("line-0", 20.0)
        },
        LayoutBlock {
            id: "paragraph-line-1".into(),
            flow_id: Some("paragraph".into()),
            flow_index: Some(1),
            flow_count: Some(2),
            minimum_fragments_per_page: Some(1),
            ..block("line-1", 20.0)
        },
    ];

    let error = layout_document(&request(fragments)).expect_err("invalid flow break");
    assert_eq!(error.code, "office.kernel.flow_break_invalid");
}

#[test]
fn rejects_non_finite_or_impossible_page_metrics() {
    let mut invalid = request(vec![block("one", 10.0)]);
    invalid.page.margin_top = invalid.page.height;
    let error = layout_document(&invalid).expect_err("invalid page");
    assert_eq!(error.code, "office.kernel.page_height_invalid");
}

#[test]
fn rejects_duplicate_block_ids() {
    let error = layout_document(&request(vec![block("same", 10.0), block("same", 20.0)]))
        .expect_err("duplicate block ID");
    assert_eq!(error.code, "office.kernel.block_id_duplicate");
}
