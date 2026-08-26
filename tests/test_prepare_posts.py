import prepare_posts


def test_prepare_posts_exposes_schedule_generator() -> None:
    assert callable(prepare_posts.generate_docs)
